import re
from datetime import datetime


SAFE_GLOBALS = {
    "__builtins__": {
        "True": True, "False": False, "None": None,
        "abs": abs, "int": int, "float": float,
        "str": str, "len": len, "min": min, "max": max,
        "round": round, "sum": sum,
    }
}


def _val(row, key):
    v = row.get(key, "")
    return str(v) if v is not None else ""


def op_exact(row, params):
    return _val(row, params.get("source_col", ""))


def op_constant(row, params):
    return str(params.get("value", ""))


def op_split(row, params):
    src = _val(row, params.get("source_col", ""))
    delim = params.get("delimiter", ",")
    idx = params.get("index", 0)
    parts = src.split(delim)
    if idx < len(parts):
        return parts[idx].strip()
    return ""


def op_concat(row, params):
    cols = params.get("source_cols", [])
    sep = params.get("separator", " ")
    vals = [_val(row, c) for c in cols]
    return sep.join(vals)


def op_condition(row, params):
    for rule in params.get("rules", []):
        val = _val(row, rule.get("source_col", ""))
        operator = rule.get("operator", "==")
        target = str(rule.get("value", ""))
        matched = False
        try:
            if operator == "==":
                matched = val == target
            elif operator == "!=":
                matched = val != target
            elif operator == "contains":
                matched = target in val
            elif operator == "starts_with":
                matched = val.startswith(target)
            elif operator == "ends_with":
                matched = val.endswith(target)
            elif operator == ">":
                matched = float(val) > float(target)
            elif operator == "<":
                matched = float(val) < float(target)
            elif operator == ">=":
                matched = float(val) >= float(target)
            elif operator == "<=":
                matched = float(val) <= float(target)
            elif operator == "is_empty":
                matched = val.strip() == ""
        except (ValueError, TypeError):
            continue
        if matched:
            return str(rule.get("output", ""))
    return str(params.get("default", ""))


def op_lookup(row, params):
    val = _val(row, params.get("source_col", ""))
    mapping = params.get("mapping", {})
    return str(mapping.get(val, val))


def op_regex(row, params):
    src = _val(row, params.get("source_col", ""))
    pattern = params.get("pattern", "")
    group = params.get("group", 0)
    try:
        m = re.search(pattern, src)
        if m:
            return m.group(group)
        return ""
    except re.error:
        return ""


def op_formula(row, params):
    expr = params.get("expression", "")
    if not expr.strip():
        return ""
    try:
        result = eval(expr, SAFE_GLOBALS, {"row": row})
        return str(result) if result is not None else ""
    except Exception:
        return ""


def op_date_format(row, params):
    src = _val(row, params.get("source_col", ""))
    if not src:
        return ""
    in_fmt = params.get("input_format", "%Y-%m-%d")
    out_fmt = params.get("output_format", "%d/%m/%Y")
    try:
        dt = datetime.strptime(src, in_fmt)
        return dt.strftime(out_fmt)
    except (ValueError, TypeError):
        return src


def op_case(row, params):
    src = _val(row, params.get("source_col", ""))
    case_type = params.get("case_type", "upper")
    if case_type == "upper":
        return src.upper()
    elif case_type == "lower":
        return src.lower()
    elif case_type == "title":
        return src.title()
    return src


def op_strip(row, params):
    return _val(row, params.get("source_col", "")).strip()


def op_skip(row, params):
    return ""


def op_count(row, params):
    cols = params.get("source_cols", [])
    count = sum(1 for c in cols if _val(row, c).strip() != "")
    return str(count)


OPERATIONS = {
    "exact": op_exact,
    "constant": op_constant,
    "split": op_split,
    "concat": op_concat,
    "condition": op_condition,
    "lookup": op_lookup,
    "regex": op_regex,
    "formula": op_formula,
    "date_format": op_date_format,
    "case": op_case,
    "strip": op_strip,
    "skip": op_skip,
    "count": op_count,
}


def apply_mapping(source_rows, mappings):
    result_rows = []
    errors_rows = []
    for row in source_rows:
        out = {}
        err = {}
        for m in mappings:
            tcol = m["target_col"]
            op = m["op"]
            params = m.get("params", {})
            handler = OPERATIONS.get(op, op_skip)
            try:
                out[tcol] = handler(row, params)
                err[tcol] = ""
            except Exception as e:
                out[tcol] = ""
                err[tcol] = str(e)
        result_rows.append(out)
        errors_rows.append(err)
    return result_rows, errors_rows


def get_supported_operations():
    return sorted(OPERATIONS.keys())
