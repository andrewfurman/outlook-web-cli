//! Small helpers that give request bodies the same loose meaning the original
//! Node.js version gave them: `--limit 20` arrives as the string "20", `--unread`
//! as "true", and `--json` can pass real numbers, booleans and arrays. These
//! follow JavaScript's `Number()`, `String()`, truthiness, `||` and `??`.

use serde_json::{Map, Value};

pub type Body = Map<String, Value>;

/// JavaScript truthiness.
pub fn truthy(v: &Value) -> bool {
    match v {
        Value::Null => false,
        Value::Bool(b) => *b,
        Value::Number(n) => n.as_f64().map(|f| f != 0.0 && !f.is_nan()).unwrap_or(true),
        Value::String(s) => !s.is_empty(),
        Value::Array(_) | Value::Object(_) => true,
    }
}

/// `body.a || body.b || ...`: the first truthy value.
pub fn or(body: &Body, keys: &[&str]) -> Option<Value> {
    keys.iter()
        .filter_map(|k| body.get(*k))
        .find(|v| truthy(v))
        .cloned()
}

/// `body.a ?? body.b ?? ...`: the first value that is present and not null.
pub fn nn(body: &Body, keys: &[&str]) -> Option<Value> {
    keys.iter()
        .filter_map(|k| body.get(*k))
        .find(|v| !v.is_null())
        .cloned()
}

/// A trimmed string, or "" for anything that is not a non-blank string.
pub fn normalize_string(v: Option<&Value>) -> String {
    match v {
        Some(Value::String(s)) => s.trim().to_string(),
        _ => String::new(),
    }
}

/// JavaScript `Number(v)`; `None` (undefined) is NaN.
pub fn number(v: Option<&Value>) -> f64 {
    match v {
        None => f64::NAN,
        Some(Value::Null) => 0.0,
        Some(Value::Bool(b)) => f64::from(u8::from(*b)),
        Some(Value::Number(n)) => n.as_f64().unwrap_or(f64::NAN),
        Some(Value::String(s)) => string_to_number(s),
        Some(Value::Array(a)) => match a.len() {
            0 => 0.0,
            1 => number(Some(&Value::String(string(&a[0])))),
            _ => f64::NAN,
        },
        Some(Value::Object(_)) => f64::NAN,
    }
}

fn string_to_number(s: &str) -> f64 {
    let t = s.trim();
    if t.is_empty() {
        return 0.0;
    }
    let lower = t.to_ascii_lowercase();
    let radix = |p: &str, r: u32| {
        i64::from_str_radix(p, r)
            .map(|n| n as f64)
            .unwrap_or(f64::NAN)
    };
    if let Some(h) = lower.strip_prefix("0x") {
        return radix(h, 16);
    }
    if let Some(o) = lower.strip_prefix("0o") {
        return radix(o, 8);
    }
    if let Some(b) = lower.strip_prefix("0b") {
        return radix(b, 2);
    }
    match t {
        "Infinity" | "+Infinity" => return f64::INFINITY,
        "-Infinity" => return f64::NEG_INFINITY,
        _ => {}
    }
    if lower.contains("inf") || lower.contains("nan") {
        return f64::NAN;
    }
    t.parse::<f64>().unwrap_or(f64::NAN)
}

/// JavaScript `String(v)`.
pub fn string(v: &Value) -> String {
    match v {
        Value::Null => "null".into(),
        Value::Bool(b) => b.to_string(),
        Value::Number(n) => match n.as_f64() {
            Some(f) => format_number(f),
            None => n.to_string(),
        },
        Value::String(s) => s.clone(),
        Value::Array(a) => a
            .iter()
            .map(|x| {
                if x.is_null() {
                    String::new()
                } else {
                    string(x)
                }
            })
            .collect::<Vec<_>>()
            .join(","),
        Value::Object(_) => "[object Object]".into(),
    }
}

pub fn format_number(f: f64) -> String {
    if f.is_nan() {
        "NaN".into()
    } else if f.is_infinite() {
        if f > 0.0 { "Infinity" } else { "-Infinity" }.into()
    } else if f.fract() == 0.0 && f.abs() < 1e21 {
        format!("{}", f as i64)
    } else {
        format!("{f}")
    }
}

/// A JSON number the way `JSON.stringify` writes it: whole numbers without ".0",
/// and null for NaN or infinity.
pub fn num(f: f64) -> Value {
    if !f.is_finite() {
        Value::Null
    } else if f.fract() == 0.0 && f.abs() < 9.0e15 {
        Value::from(f as i64)
    } else {
        serde_json::Number::from_f64(f)
            .map(Value::Number)
            .unwrap_or(Value::Null)
    }
}

/// `String(v || "")`.
pub fn str_or_empty(v: Option<&Value>) -> String {
    match v {
        Some(x) if truthy(x) => string(x),
        _ => String::new(),
    }
}

/// `clampInteger` from the Node.js version: Number(), truncate, clamp, else fallback.
pub fn clamp_integer(v: Option<&Value>, min: i64, max: i64, fallback: i64) -> i64 {
    let f = number(v);
    if !f.is_finite() {
        return fallback;
    }
    (f.trunc() as i64).clamp(min, max)
}

/// `toBoolean` from the Node.js version.
pub fn to_boolean(v: Option<&Value>, fallback: bool) -> bool {
    match v {
        Some(Value::Bool(b)) => *b,
        Some(Value::String(s)) if s == "true" || s == "1" => true,
        Some(Value::String(s)) if s == "false" || s == "0" => false,
        Some(Value::Number(n)) if n.as_f64() == Some(1.0) => true,
        Some(Value::Number(n)) if n.as_f64() == Some(0.0) => false,
        _ => fallback,
    }
}

/// Present and not null or "" (the `x !== undefined && x !== null && x !== ""` test).
pub fn given(v: Option<&Value>) -> bool {
    !matches!(v, None | Some(Value::Null)) && v != Some(&Value::String(String::new()))
}

/// Truncates to `n` characters, ending with "..." when cut.
pub fn trim_to(s: &str, n: usize) -> String {
    if s.chars().count() <= n {
        return s.to_string();
    }
    let keep: String = s.chars().take(n.saturating_sub(3)).collect();
    format!("{keep}...")
}

pub fn first_line(s: &str) -> String {
    s.split('\n')
        .next()
        .unwrap_or("")
        .trim_end_matches('\r')
        .chars()
        .take(200)
        .collect()
}

pub fn plural(n: usize) -> &'static str {
    if n == 1 {
        ""
    } else {
        "s"
    }
}

/// Follows a path of object keys.
pub fn get<'a>(v: &'a Value, path: &[&str]) -> Option<&'a Value> {
    let mut cur = v;
    for k in path {
        cur = cur.get(*k)?;
    }
    Some(cur)
}

/// `String(v?.a?.b || "")`.
pub fn get_str(v: &Value, path: &[&str]) -> String {
    str_or_empty(get(v, path))
}

/// Sets `key`, or removes it when `value` is None (JSON.stringify drops undefined).
pub fn set_opt(map: &mut Body, key: &str, value: Option<Value>) {
    match value {
        Some(v) => {
            map.insert(key.to_string(), v);
        }
        None => {
            map.shift_remove(key);
        }
    }
}

/// Copies every field of `from` into `into`, like `{...into, ...from}`.
pub fn assign(into: &mut Body, from: Value) {
    if let Value::Object(m) = from {
        for (k, v) in m {
            into.insert(k, v);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn numbers_follow_javascript() {
        assert_eq!(number(Some(&json!("20"))), 20.0);
        assert_eq!(number(Some(&json!(" 7 "))), 7.0);
        assert_eq!(number(Some(&json!(""))), 0.0);
        assert!(number(Some(&json!("abc"))).is_nan());
        assert!(number(None).is_nan());
        assert_eq!(clamp_integer(Some(&json!("200")), 1, 50, 20), 50);
        assert_eq!(clamp_integer(None, 1, 50, 20), 20);
        assert_eq!(num(23.4), json!(23.4));
        assert_eq!(num(3.0), json!(3));
        assert_eq!(trim_to("abcdef", 5), "ab...");
    }
}
