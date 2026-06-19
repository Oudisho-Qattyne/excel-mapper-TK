from dataclasses import dataclass, field


@dataclass
class AppState:
    source_rows: list[dict] = field(default_factory=list)
    source_columns: list[str] = field(default_factory=list)
    source_file_name: str = ""
    source_sheet_names: list[str] = field(default_factory=list)
    source_selected_sheet: str = ""

    target_rows: list[dict] = field(default_factory=list)
    target_columns: list[str] = field(default_factory=list)
    target_file_name: str = ""
    target_sheet_names: list[str] = field(default_factory=list)
    target_selected_sheet: str = ""

    mappings: list[dict] = field(default_factory=list)
    result_rows: list[dict] = field(default_factory=list)
    result_errors: list[dict] = field(default_factory=list)
