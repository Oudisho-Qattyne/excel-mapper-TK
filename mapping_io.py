import json
import os

VALID_OPS = {
    "exact", "constant", "split", "concat", "condition",
    "lookup", "regex", "formula", "date_format", "case",
    "strip", "skip", "count",
}


def save_mapping(mappings, path):
    config = {"mappings": mappings}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def load_mapping(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Mapping file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        config = json.load(f)
    errors = validate_mapping(config)
    if errors:
        raise ValueError("Invalid mapping config:\n" + "\n".join(errors))
    return config["mappings"]


def validate_mapping(config):
    errors = []
    if not isinstance(config, dict):
        return ["Root must be a dict with a 'mappings' list"]
    mappings = config.get("mappings", [])
    if not isinstance(mappings, list):
        return ["'mappings' must be a list"]
    for i, m in enumerate(mappings):
        if "target_col" not in m:
            errors.append(f"Mapping #{i}: missing 'target_col'")
        if "op" not in m:
            errors.append(f"Mapping #{i}: missing 'op'")
        elif m["op"] not in VALID_OPS:
            errors.append(f"Mapping #{i}: unknown op '{m['op']}'")
        if "params" not in m:
            errors.append(f"Mapping #{i}: missing 'params'")
        elif not isinstance(m["params"], dict):
            errors.append(f"Mapping #{i}: 'params' must be a dict")
    return errors
