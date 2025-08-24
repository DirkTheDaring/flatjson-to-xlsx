use calamine::{open_workbook_auto, DataType, Reader};
use regex::Regex;
use serde::Deserialize;
use serde_json::{Map as JsonMap, Value};
use std::cmp::Ordering;
use std::collections::{BTreeSet, HashMap, HashSet};
use std::env;
use std::fs;
use std::io::{self, Read};
use std::path::Path;
use umya_spreadsheet as umya;

const VERSION: &str = "0.7.0";

#[derive(Debug, Default, Deserialize)]
struct ConfigFile {
    out: Option<String>,
    sheet: Option<String>,
    ndjson: Option<bool>,
    pk: Option<Vec<String>>,

    // include filters
    include: Option<Vec<String>>,
    include_regex: Option<Vec<String>>,
    include_substr: Option<Vec<String>>,

    // ordering
    order: Option<Vec<String>>,
    order_regex: Option<Vec<String>>,
    order_substr: Option<Vec<String>>,
    order_rest: Option<String>, // existing|alpha|none

    // whether PKs are forced to appear first (default true)
    pk_first: Option<bool>,

    // NEW: per-column hyperlink bases (exact column names)
    #[serde(default)]
    hyperlink: HashMap<String, String>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();

    if has_flag(&args, "--help") || has_flag(&args, "-h") {
        print_help(&args[0]);
        return Ok(());
    }
    if has_flag(&args, "--version") || has_flag(&args, "-V") {
        println!("{} {}", args[0], VERSION);
        return Ok(());
    }

    // -------- config (optional)
    let cfg_path = get_arg_value(&args, "--config").or_else(|| get_arg_value(&args, "-c"));
    let cfg = if let Some(path) = cfg_path.as_deref() {
        Some(load_config(path)?)
    } else {
        None
    };

    // -------- resolve options
    let out_path = get_arg_value(&args, "--out")
        .or_else(|| get_arg_value(&args, "-o"))
        .or_else(|| cfg.as_ref().and_then(|c| c.out.clone()))
        .expect("--out <FILE.xlsx> is required (or set in config)");
    if !out_path.to_lowercase().ends_with(".xlsx") {
        eprintln!("--out must end with .xlsx");
        std::process::exit(2);
    }

    let sheet_name = get_arg_value(&args, "--sheet")
        .or_else(|| get_arg_value(&args, "-s"))
        .or_else(|| cfg.as_ref().and_then(|c| c.sheet.clone()))
        .unwrap_or_else(|| "Sheet1".to_string());

    // Input mode precedence: --array > --ndjson > config.ndjson > false
    let force_array = has_flag(&args, "--array");
    let force_ndjson = has_flag(&args, "--ndjson");
    let mut ndjson_mode = if force_array {
        false
    } else if force_ndjson {
        true
    } else {
        cfg.as_ref().and_then(|c| c.ndjson).unwrap_or(false)
    };

    // PKs
    let pk_keys: Vec<String> = if let Some(s) = get_arg_value(&args, "--pk").or_else(|| get_arg_value(&args, "-k")) {
        split_csv(&s)
    } else if let Some(c) = &cfg {
        c.pk.clone().unwrap_or_default()
    } else {
        Vec::new()
    };

    // pk_first toggle (default true). CLI supports --pk-first / --no-pk-first
    let pk_first = if has_flag(&args, "--no-pk-first") {
        false
    } else if has_flag(&args, "--pk-first") {
        true
    } else {
        cfg.as_ref().and_then(|c| c.pk_first).unwrap_or(true)
    };

    // ---------------- include filters ----------------
    let include_flag_exact_cli = has_flag(&args, "--include") || has_flag(&args, "-i");
    let mut include_exact: HashSet<String> = cfg
        .as_ref()
        .and_then(|c| c.include.clone())
        .unwrap_or_default()
        .into_iter()
        .collect();
    if let Some(s) = get_arg_value(&args, "--include").or_else(|| get_arg_value(&args, "-i")) {
        include_exact.extend(split_csv(&s).into_iter());
    }

    let include_flag_regex_cli = has_flag(&args, "--include-regex");
    let mut include_regexes: Vec<Regex> = Vec::new();
    if let Some(c) = &cfg {
        for pat in c.include_regex.clone().unwrap_or_default() {
            include_regexes.push(
                Regex::new(&format!("^(?:{})$", pat))
                    .map_err(|e| format!("Invalid regex in config `{}`: {}", pat, e))?,
            );
        }
    }
    if let Some(s) = get_arg_value(&args, "--include-regex") {
        for pat in split_csv(&s).into_iter().filter(|p| !p.is_empty()) {
            include_regexes.push(
                Regex::new(&format!("^(?:{})$", pat))
                    .map_err(|e| format!("Invalid regex `{}`: {}", pat, e))?,
            );
        }
    }

    let include_flag_substr_cli = has_flag(&args, "--include-substr");
    let mut include_substrs: Vec<String> = cfg
        .as_ref()
        .and_then(|c| c.include_substr.clone())
        .unwrap_or_default();
    if let Some(s) = get_arg_value(&args, "--include-substr") {
        include_substrs.extend(split_csv(&s));
    }

    let include_active = include_flag_exact_cli
        || include_flag_regex_cli
        || include_flag_substr_cli
        || !include_exact.is_empty()
        || !include_regexes.is_empty()
        || !include_substrs.is_empty();

    // ---------------- ordering controls ----------------
    let order_exact: Vec<String> = if let Some(s) = get_arg_value(&args, "--order") {
        split_csv(&s)
    } else if let Some(c) = &cfg {
        c.order.clone().unwrap_or_default()
    } else {
        Vec::new()
    };

    let order_regex_raw: Vec<String> = if let Some(s) = get_arg_value(&args, "--order-regex") {
        split_csv(&s)
    } else if let Some(c) = &cfg {
        c.order_regex.clone().unwrap_or_default()
    } else {
        Vec::new()
    };
    let order_regex: Vec<Regex> = order_regex_raw
        .iter()
        .map(|pat| Regex::new(&format!("^(?:{})$", pat))
            .map_err(|e| format!("Invalid --order-regex `{}`: {}", pat, e)))
        .collect::<Result<_, _>>()?;

    let order_substr: Vec<String> = if let Some(s) = get_arg_value(&args, "--order-substr") {
        split_csv(&s)
    } else if let Some(c) = &cfg {
        c.order_substr.clone().unwrap_or_default()
    } else {
        Vec::new()
    };

    let order_rest: String = if let Some(s) = get_arg_value(&args, "--order-rest") {
        s
    } else if let Some(c) = &cfg {
        c.order_rest.clone().unwrap_or_else(|| "existing".to_string())
    } else {
        "existing".to_string()
    };
    let order_rest_mode = order_rest.to_lowercase();

    // ------------- hyperlinks (config + CLI) -------------
    let mut hyperlink_map: HashMap<String, String> = cfg
        .as_ref()
        .map(|c| c.hyperlink.clone())
        .unwrap_or_default();

    if let Some(link_arg) = get_arg_value(&args, "--link") {
        // format: col=BASE[,col2=BASE2,...]
        for part in split_csv(&link_arg) {
            if let Some((k, v)) = split_once_eq(&part) {
                hyperlink_map.insert(k.to_string(), v.to_string());
            } else {
                eprintln!("Ignoring malformed --link mapping: `{}` (expected col=BASE)", part);
            }
        }
    }

    // -------- read stdin --------
    let mut input = String::new();
    io::stdin().read_to_string(&mut input)?;

    // If in NDJSON mode but the payload clearly looks like a JSON array, switch modes.
    let looks_like_array = input.trim_start().starts_with('[');
    if ndjson_mode && looks_like_array {
        eprintln!("Note: input looks like a JSON array; overriding NDJSON and parsing as array.");
        ndjson_mode = false;
    }

    let mut new_rows: Vec<HashMap<String, Value>> = if ndjson_mode {
        parse_ndjson(&input)?
    } else {
        parse_json_array(&input)?
    };

    // -------- existing workbook values --------
    let out_exists = Path::new(&out_path).exists();
    let (mut existing_headers, mut existing_rows) = if out_exists {
        match read_existing_xlsx_values(&out_path, &sheet_name) {
            Ok(data) => data,
            Err(e) => {
                eprintln!(
                    "Warning: couldn't read existing workbook `{}` ({}). Recreating sheet content.",
                    out_path, e
                );
                (Vec::new(), Vec::new())
            }
        }
    } else {
        (Vec::new(), Vec::new())
    };

    // -------- merge by PK --------
    if !pk_keys.is_empty() {
        let mut index: HashMap<String, usize> = HashMap::new();
        for (i, row) in existing_rows.iter().enumerate() {
            if let Some(key) = composite_pk(row, &pk_keys) {
                index.insert(key, i);
            }
        }
        for row in new_rows.drain(..) {
            if let Some(key) = composite_pk(&row, &pk_keys) {
                if let Some(&idx) = index.get(&key) {
                    existing_rows[idx] = row;
                } else {
                    index.insert(key, existing_rows.len());
                    existing_rows.push(row);
                }
            } else {
                existing_rows.push(row);
            }
        }
    } else {
        existing_rows.extend(new_rows.drain(..));
    }

    // -------- union of keys --------
    let mut all_keys: BTreeSet<String> = BTreeSet::new();
    for r in &existing_rows {
        for k in r.keys() {
            all_keys.insert(k.clone());
        }
    }

    // clean empty header cells
    existing_headers.retain(|h| !h.trim().is_empty());

    // inclusion predicate
    let pk_set: HashSet<&str> = pk_keys.iter().map(|s| s.as_str()).collect();
    let key_allowed = |key: &str| -> bool {
        if pk_set.contains(key) {
            return true; // PKs always included even if include filters present
        }
        if include_active {
            if include_exact.contains(key) {
                return true;
            }
            if !include_regexes.is_empty() && include_regexes.iter().any(|re| re.is_match(key)) {
                return true;
            }
            if !include_substrs.is_empty() && include_substrs.iter().any(|sub| key.contains(sub)) {
                return true;
            }
            false
        } else {
            true
        }
    };

    // universe: existing headers (filtered) + remaining keys (natural sorted), all filtered by key_allowed
    let universe_existing: Vec<String> = existing_headers
        .iter()
        .filter(|h| !pk_set.contains(h.as_str()) && key_allowed(h))
        .cloned()
        .collect();

    let mut remaining_from_all: Vec<String> = all_keys
        .iter()
        .filter(|k| !pk_set.contains(k.as_str()) && key_allowed(k))
        .cloned()
        .collect();
    remaining_from_all.retain(|k| !universe_existing.contains(k));
    remaining_from_all.sort_by(natural_cmp);

    let mut universe: Vec<String> = Vec::new();
    universe.extend(universe_existing.into_iter());
    universe.extend(remaining_from_all.into_iter());

    // ---------- build final columns ----------
    let mut columns: Vec<String> = Vec::new();
    let mut seen: HashSet<String> = HashSet::new();

    // 1) PKs first (if configured)
    if pk_first {
        for pk in &pk_keys {
            push_unique(&mut columns, &mut seen, pk.clone());
        }
    }

    // 2) ordered groups
    add_exact(&mut columns, &mut seen, &order_exact, &key_allowed);
    add_regex(&mut columns, &mut seen, &universe, &order_regex);
    add_substr(&mut columns, &mut seen, &universe, &order_substr);

    // 3) remainder
    match order_rest_mode.as_str() {
        "none" => {}
        "alpha" => {
            let mut rest: Vec<String> = universe.into_iter().filter(|k| !seen.contains(k)).collect();
            rest.sort_by(natural_cmp);
            for k in rest {
                push_unique(&mut columns, &mut seen, k);
            }
        }
        _ /* existing */ => {
            let rest: Vec<String> = universe.into_iter().filter(|k| !seen.contains(k)).collect();
            for k in rest {
                push_unique(&mut columns, &mut seen, k);
            }
        }
    }

    // 4) ensure PKs exist even if pk_first=false and not covered above
    if !pk_first {
        for pk in &pk_keys {
            if !seen.contains(pk) {
                push_unique(&mut columns, &mut seen, pk.clone());
            }
        }
    }

    // -------- write/update XLSX while preserving formatting --------
    write_xlsx_preserve(&out_path, &sheet_name, &columns, &existing_rows, &hyperlink_map)?;
    Ok(())
}

// ---------------- helpers: ordering adders ----------------

fn add_exact<F: Fn(&str) -> bool>(
    columns: &mut Vec<String>,
    seen: &mut HashSet<String>,
    names: &[String],
    key_allowed: &F,
) {
    for name in names {
        if key_allowed(name) {
            push_unique(columns, seen, name.clone());
        }
    }
}

fn add_regex(
    columns: &mut Vec<String>,
    seen: &mut HashSet<String>,
    universe: &[String],
    pats: &[Regex],
) {
    for re in pats {
        for k in universe.iter().filter(|k| re.is_match(k)) {
            push_unique(columns, seen, k.clone());
        }
    }
}

fn add_substr(
    columns: &mut Vec<String>,
    seen: &mut HashSet<String>,
    universe: &[String],
    subs: &[String],
) {
    for sub in subs {
        for k in universe.iter().filter(|k| k.contains(sub)) {
            push_unique(columns, seen, k.clone());
        }
    }
}

// ---------------- Parsing ----------------

fn parse_json_array(input: &str) -> Result<Vec<HashMap<String, Value>>, Box<dyn std::error::Error>> {
    let v: Value = serde_json::from_str(input)?;
    match v {
        Value::Array(arr) => arr.into_iter().map(value_to_rowmap).collect(),
        Value::Object(obj) => Ok(vec![object_to_rowmap(obj)]),
        _ => Err("Expected a JSON array of objects or a single object".into()),
    }
}

fn parse_ndjson(input: &str) -> Result<Vec<HashMap<String, Value>>, Box<dyn std::error::Error>> {
    let mut rows = Vec::new();
    for (lineno, line) in input.lines().enumerate() {
        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        let v: Value = serde_json::from_str(line)
            .map_err(|e| format!("Invalid JSON on line {}: {}", lineno + 1, e))?;
        rows.push(value_to_rowmap(v)?);
    }
    Ok(rows)
}

fn value_to_rowmap(v: Value) -> Result<HashMap<String, Value>, Box<dyn std::error::Error>> {
    match v {
        Value::Object(obj) => Ok(object_to_rowmap(obj)),
        _ => Err("Each record must be a JSON object (already flattened)".into()),
    }
}

fn object_to_rowmap(obj: JsonMap<String, Value>) -> HashMap<String, Value> {
    obj.into_iter().collect()
}

// ---------------- Read existing values (calamine) ----------------

fn read_existing_xlsx_values(
    path: &str,
    sheet_name: &str,
) -> Result<(Vec<String>, Vec<HashMap<String, Value>>), Box<dyn std::error::Error>> {
    let mut wb = open_workbook_auto(path)?;
    let maybe_range = wb.worksheet_range(sheet_name);

    let range = match maybe_range {
        Some(Ok(r)) => r,
        Some(Err(e)) => return Err(Box::<dyn std::error::Error>::from(e)),
        None => return Ok((Vec::new(), Vec::new())),
    };

    let mut rows_iter = range.rows();
    let header_cells = match rows_iter.next() {
        Some(r) => r.to_vec(),
        None => return Ok((Vec::new(), Vec::new())),
    };

    let headers: Vec<String> = header_cells.iter().map(cell_to_string).collect();

    let mut rows: Vec<HashMap<String, Value>> = Vec::new();
    for r in rows_iter {
        let mut map = HashMap::new();
        for (i, cell) in r.iter().enumerate() {
            if let Some(col) = headers.get(i) {
                if col.trim().is_empty() {
                    continue;
                }
                let s = cell_to_string(cell);
                if !s.is_empty() {
                    map.insert(col.clone(), Value::String(s));
                } else {
                    map.insert(col.clone(), Value::Null);
                }
            }
        }
        if map.values().any(|v| !v.is_null()) {
            rows.push(map);
        }
    }

    Ok((headers, rows))
}

fn cell_to_string(cell: &DataType) -> String {
    match cell {
        DataType::Empty => String::new(),
        DataType::String(s) => s.to_string(),
        DataType::Float(f) => {
            if f.fract() == 0.0 {
                format!("{}", *f as i64)
            } else {
                f.to_string()
            }
        }
        DataType::Int(i) => i.to_string(),
        DataType::Bool(b) => b.to_string(),
        DataType::Error(e) => format!("ERR:{:?}", e),
        DataType::DateTime(v) => v.to_string(),
        DataType::Duration(v) => v.to_string(),
        DataType::DateTimeIso(s) => s.clone(),
        DataType::DurationIso(s) => s.clone(),
    }
}

// ---------------- PK handling ----------------

fn composite_pk(row: &HashMap<String, Value>, pk_cols: &[String]) -> Option<String> {
    let mut parts: Vec<String> = Vec::with_capacity(pk_cols.len());
    for c in pk_cols {
        match row.get(c) {
            Some(Value::String(s)) => parts.push(s.clone()),
            Some(Value::Number(n)) => parts.push(n.to_string()),
            Some(Value::Bool(b)) => parts.push(b.to_string()),
            Some(Value::Null) | None => return None,
            Some(other) => parts.push(other.to_string()),
        }
    }
    Some(parts.join("\u{1F}"))
}

// ---------------- Write XLSX while preserving formatting ----------------

fn write_xlsx_preserve(
    out_path: &str,
    sheet_name: &str,
    columns: &[String],
    rows: &[HashMap<String, Value>],
    hyperlink_map: &HashMap<String, String>,
) -> Result<(), Box<dyn std::error::Error>> {
    // Open existing workbook or create a new one
    let mut book = if Path::new(out_path).exists() {
        umya::reader::xlsx::read(Path::new(out_path))?
    } else {
        umya::new_file()
    };

    // Ensure sheet exists (create or rename default)
    if book.get_sheet_by_name(sheet_name).is_none() {
        if sheet_name != "Sheet1" {
            if let Some(ws1) = book.get_sheet_by_name_mut("Sheet1") {
                ws1.set_name(sheet_name);
            } else {
                let _ = book.new_sheet(sheet_name);
            }
        } else {
            let _ = book.new_sheet(sheet_name);
        }
    }

    // Now we can safely get it mutably
    let ws = book
        .get_sheet_by_name_mut(sheet_name)
        .expect("sheet must exist");

    // Header row (keeps existing styles)
    for (c_idx, col_name) in columns.iter().enumerate() {
        let col = (c_idx as u32) + 1;
        ws.get_cell_mut((col, 1)).set_value(col_name);
    }

    // Data rows (starting at row 2) — preserves formatting of those cells
    for (r_idx, rowmap) in rows.iter().enumerate() {
        let row_num = (r_idx as u32) + 2;
        for (c_idx, key) in columns.iter().enumerate() {
            let col = (c_idx as u32) + 1;
            let cell = ws.get_cell_mut((col, row_num));

            if let Some(v) = rowmap.get(key) {
                // If this column is mapped to a hyperlink base, write a HYPERLINK formula
                if let Some(base) = hyperlink_map.get(key) {
                    // Build display text from the value
                    let text = match v {
                        Value::Null => "".to_string(),
                        Value::Bool(b) => b.to_string(),
                        Value::Number(n) => n.to_string(),
                        Value::String(s) => s.clone(),
                        other => other.to_string(),
                    };
                    if !text.is_empty() {
                        let url = format!("{}{}", base, &text);
                        let f = format!(
                            "HYPERLINK(\"{}\",\"{}\")",
                            xl_quote_escape(&url),
                            xl_quote_escape(&text)
                        );
                        cell.set_formula(&f);
                        continue;
                    } else {
                        // empty text => write empty value (no link)
                        cell.set_value("");
                        continue;
                    }
                }

                // Normal write for non-hyperlink columns
                match v {
                    Value::Null => {
                        cell.set_value("");
                    }
                    Value::Bool(b) => {
                        cell.set_value_bool(*b);
                    }
                    Value::Number(n) => {
                        if let Some(f) = n.as_f64() {
                            cell.set_value_number(f);
                        } else {
                            cell.set_value(n.to_string());
                        }
                    }
                    Value::String(s) => {
                        cell.set_value(s);
                    }
                    other => {
                        cell.set_value(other.to_string());
                    }
                }
            }
        }
    }

    // Save back to same file (styles remain intact)
    umya::writer::xlsx::write(&book, Path::new(out_path))?;
    Ok(())
}

// ---------------- misc helpers ----------------

fn load_config(path: &str) -> Result<ConfigFile, Box<dyn std::error::Error>> {
    let txt = fs::read_to_string(path)?;
    let cfg: ConfigFile = toml::from_str(&txt)?;
    Ok(cfg)
}

fn print_help(program: &str) {
    println!("Usage:");
    println!("  {program} --out OUT.xlsx [--sheet Sheet1] [--pk col1,col2,...] \\");
    println!("            [--array | --ndjson] \\");
    println!("            [--include name1,name2,...] [--include-regex r1,r2,...] [--include-substr s1,s2,...] \\");
    println!("            [--order n1,n2,...] [--order-regex r1,r2,...] [--order-substr s1,s2,...] [--order-rest existing|alpha|none] \\");
    println!("            [--pk-first | --no-pk-first] [--link col=BASE[,col2=BASE2,...]] [--config file.toml] < input.json");
    println!();
    println!("Notes:");
    println!("  • Existing XLSX is updated in-place; formatting is preserved.");
    println!("  • If NDJSON is set but input starts with '[', the tool switches to array mode.");
    println!("  • Inclusion is ACTIVE if any include list is present (exact/regex/substr).");
    println!("  • Column order: (PKs if pk_first) -> ordered groups -> remainder (order-rest).");
    println!("  • --link/ [hyperlink] will write a HYPERLINK formula so the cell shows only the value but is clickable.");
}

fn has_flag(args: &[String], flag: &str) -> bool {
    args.iter().any(|a| a == flag)
}

fn get_arg_value(args: &[String], flag: &str) -> Option<String> {
    args.iter()
        .position(|a| a == flag)
        .and_then(|i| args.get(i + 1))
        .cloned()
}

fn split_csv(s: &str) -> Vec<String> {
    s.split(',')
        .map(|t| t.trim().to_string())
        .filter(|t| !t.is_empty())
        .collect::<Vec<_>>()
}

fn split_once_eq(s: &str) -> Option<(&str, &str)> {
    let mut it = s.splitn(2, '=');
    let k = it.next()?;
    let v = it.next()?;
    Some((k.trim(), v.trim()))
}

fn push_unique(vec: &mut Vec<String>, seen: &mut HashSet<String>, k: String) {
    if seen.insert(k.clone()) {
        vec.push(k);
    }
}

// Natural sort so ...comments.2... < ...comments.10...
fn natural_cmp(a: &String, b: &String) -> Ordering {
    natural_cmp_str(a.as_str(), b.as_str())
}
fn natural_cmp_str(a: &str, b: &str) -> Ordering {
    let pa = natural_parts(a);
    let pb = natural_parts(b);
    let mut i = 0usize;
    while i < pa.len() && i < pb.len() {
        match (&pa[i], &pb[i]) {
            (NatPart::Num(x), NatPart::Num(y)) => match x.cmp(y) {
                Ordering::Equal => {}
                ord => return ord,
            },
            (NatPart::Txt(x), NatPart::Txt(y)) => match x.cmp(y) {
                Ordering::Equal => {}
                ord => return ord,
            },
            (NatPart::Num(_), NatPart::Txt(_)) => return Ordering::Less,
            (NatPart::Txt(_), NatPart::Num(_)) => return Ordering::Greater,
        }
        i += 1;
    }
    pa.len().cmp(&pb.len())
}
#[derive(Debug)]
enum NatPart {
    Txt(String),
    Num(u64),
}
fn natural_parts(s: &str) -> Vec<NatPart> {
    let mut out = Vec::new();
    let mut buf = String::new();
    let mut in_num = false;
    for ch in s.chars() {
        if ch.is_ascii_digit() {
            if !in_num && !buf.is_empty() {
                out.push(NatPart::Txt(buf.clone()));
                buf.clear();
            }
            in_num = true;
            buf.push(ch);
        } else {
            if in_num {
                let n = buf.parse::<u64>().unwrap_or(0);
                out.push(NatPart::Num(n));
                buf.clear();
            }
            in_num = false;
            buf.push(ch);
        }
    }
    if !buf.is_empty() {
        if in_num {
            let n = buf.parse::<u64>().unwrap_or(0);
            out.push(NatPart::Num(n));
        } else {
            out.push(NatPart::Txt(buf));
        }
    }
    out
}

// Excel formula quote-escape: " -> ""
fn xl_quote_escape(s: &str) -> String {
    s.replace('"', "\"\"")
}
