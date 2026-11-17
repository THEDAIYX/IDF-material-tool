"""IDF material extractor.

Reads an Isogen/AVEVA IDF file and writes a material take-off workbook.
The GUI supports a new workbook or an existing custom template.  The same
workflow is available from the command line for batch use.
"""

from __future__ import annotations

import argparse
import math
import re
import shutil
import sys
from collections import defaultdict
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Optional


try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError as exc:  # pragma: no cover - shown by the GUI/CLI
    raise SystemExit(
        "缺少 openpyxl。请先执行: pip install -r requirements.txt"
    ) from exc


TYPE_NAMES = {
    30: "弯头/U形弯头",
    31: "弯头/U形弯头",
    35: "弯头",
    36: "弯头/变径弯头",
    40: "Olet/仪表三通主件",
    41: "Olet/仪表三通支管",
    42: "Olet/仪表三通支管",
    45: "三通",
    46: "三通",
    47: "三通",
    50: "四通",
    51: "四通",
    52: "四通",
    53: "四通",
    55: "同心/偏心异径管",
    60: "异径三通",
    61: "异径三通",
    62: "异径三通",
    65: "异径法兰",
    70: "三通弯头",
    71: "三通弯头",
    72: "三通弯头",
    75: "角式阀",
    76: "角式阀",
    80: "三通阀",
    81: "三通阀",
    82: "三通阀",
    85: "四通阀",
    86: "四通阀",
    87: "四通阀",
    88: "四通阀",
    90: "仪表",
    91: "三通仪表",
    92: "四通仪表",
    93: "仪表",
    95: "杂项管件",
    96: "杂项管件",
    100: "管子",
    101: "定长管",
    102: "定长管段",
    103: "变长管段",
    104: "间隙件",
    105: "法兰",
    106: "活套短节/环",
    107: "盲法兰",
    110: "垫片",
    111: "卫生级连接件",
    112: "卫生级背帽",
    113: "卫生级卡箍",
    114: "卫生级杂项件",
    115: "螺栓",
    120: "焊口",
    125: "管帽",
    126: "管箍/Elbolet",
    127: "活接头",
    130: "阀门",
    132: "疏水阀/捕集器",
    133: "疏水阀/捕集器",
    134: "安全盘/放空件",
    136: "过滤器",
    137: "过滤器",
    140: "仪表气泡",
    150: "管道支架",
}

SPECIAL_TYPES = {0, 3, 148, 149, 150, 151, 152, 153, 160, 161, 200, 201, 300, 301}
NON_PHYSICAL_TYPES = {0, 3, 148, 149, 151, 152, 153, 160, 161, 200, 201, 300, 301, 999}
KNOWN_HEADER_RECORDS = {
    -5: "生成程序",
    -6: "管线号",
    -7: "线轴前缀",
    -8: "修订号",
    -9: "项目号",
    -10: "批次/区域",
    -11: "管道规格",
    -12: "额定等级",
    -13: "管线类型",
    -14: "创建日期",
    -15: "保温规格",
    -16: "伴热规格",
    -17: "油漆规格",
    -18: "比重",
    -19: "温度",
    -22: "组件标签",
    -25: "系统等轴图名",
    -31: "设备连接",
}


@dataclass
class MaterialDef:
    index: int
    code: str
    description: str


@dataclass
class Component:
    line_no: int
    type_id: int
    start_raw: tuple[int, int, int]
    end_raw: tuple[int, int, int]
    bore: int
    item_index: int
    weight: str
    word11: str
    skey: str
    word13: str
    word14: str
    raw: str
    spec: str = ""
    spec_change_after: str = ""
    unique_id: str = ""
    tag: str = ""
    pipeline_connection: str = ""
    equipment_connection: str = ""
    user_message: str = ""
    notes: list[str] = field(default_factory=list)

    @property
    def start_mm(self) -> tuple[float, float, float]:
        return tuple(round(v / 100.0, 2) for v in self.start_raw)

    @property
    def end_mm(self) -> tuple[float, float, float]:
        return tuple(round(v / 100.0, 2) for v in self.end_raw)

    @property
    def length_mm(self) -> Optional[float]:
        if self.type_id in SPECIAL_TYPES or self.type_id == 999:
            return None
        a = self.start_mm
        b = self.end_mm
        return round(math.dist(a, b), 2)

    @property
    def type_name(self) -> str:
        if self.type_id == 149 and self.skey.upper() == "FLOW":
            return "流向箭头"
        if self.type_id == 300:
            return "公制坐标偏移"
        if self.type_id == 0:
            return "附加孔径记录"
        return TYPE_NAMES.get(self.type_id, f"类型 {self.type_id}")

    @property
    def item_class(self) -> str:
        value = _safe_int(self.word14)
        return {0: "无分类", 1: "制造", 2: "安装", 3: "海工"}.get(value % 10, "")

    @property
    def flow_direction(self) -> str:
        value = _safe_int(self.word14)
        return "与构建方向相反" if (value // 10) % 10 == 1 else "按构建方向"

    @property
    def primary_direction(self) -> str:
        value = _safe_int(self.word14)
        return {
            0: "",
            1: "北",
            2: "南",
            3: "东",
            4: "西",
            5: "上",
            6: "下",
        }.get((value // 100) % 10, "")

    @property
    def insulation_tracing(self) -> str:
        value = _safe_int(self.word14)
        return {
            0: "无",
            1: "保温",
            2: "伴热",
            3: "保温+伴热",
        }.get((value // 1000) % 10, "")

    @property
    def bom_status(self) -> str:
        value = _safe_int(self.word11)
        return {
            0: "出现在BOM并在等轴图显示料号",
            1: "不出现在BOM和等轴图",
            2: "出现在BOM和等轴图并显示在料号下方",
        }.get((value // 1_000_000) % 10, "")

    @property
    def angle_or_bolt_length(self) -> str:
        n = _safe_int(self.word13)
        if not n:
            return ""
        if self.type_id == 115:
            return f"螺栓长度 {n}"
        if self.type_id in {30, 31, 35, 36, 40, 41, 42, 45, 46, 47, 50, 51, 52, 53}:
            return f"角度 {n / 100:.2f}°"
        return str(n)


def _safe_int(value: Any) -> int:
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return 0


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def parse_spec_dimensions(description: str, bore: int | str = "") -> tuple[str, str, str]:
    """从材料描述中提取 DN、SCH 壁厚等级和 CL/PN 压力等级。"""
    text = normalize_text(description).upper()
    dn_values: list[str] = []
    dn_text = re.sub(r"[-X×]\s*[0-9]+(?:\.[0-9]+)?\s*MM\b", "", text)
    for match in re.finditer(r"\bDN\s*([0-9]+(?:\.[0-9]+)?)(?:\s*[-X×]\s*(?:DN\s*)?([0-9]+(?:\.[0-9]+)?))?", dn_text):
        dn_values.append(match.group(1))
        if match.group(2):
            dn_values.append(match.group(2))
    sch_values = [match.group(1) for match in re.finditer(r"SCH\s*([0-9]+(?:\.[0-9]+)?|STD|XS)(?=\b|X)", text)]
    wall_values = [match.group(1) for match in re.finditer(r"([0-9]+(?:\.[0-9]+)?)\s*MM", text)]
    pressure_match = re.search(r"\b(CL\s*[0-9]+|PN\s*[0-9]+|[0-9]+\s*LB)\b", text)
    dn = "x".join(f"DN{value}" for value in dn_values) if dn_values else (f"DN{bore}" if bore not in ("", None, 0) else "")
    sch = "x".join(f"SCH{value}" for value in sch_values) if sch_values else "x".join(f"{value}MM" for value in wall_values)
    pressure = normalize_text(pressure_match.group(1).replace(" ", "")) if pressure_match else ""
    return dn, sch, pressure


def bom_control_value(component: Component) -> int:
    """Word 11 的百位控制 BOM：0=进入 BOM，1=不进入 BOM。"""
    return (_safe_int(component.word11) // 100) % 10


def read_idf(path: str | Path) -> tuple[list[str], str]:
    raw = Path(path).read_bytes()
    for encoding in ("utf-8-sig", "gb18030", "cp1252", "latin-1"):
        try:
            return raw.decode(encoding).splitlines(), encoding
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace").splitlines(), "latin-1"


def _negative_match(line: str) -> Optional[tuple[int, str]]:
    match = re.match(r"^\s*(-\d+)\s?(.*)$", line)
    if not match:
        return None
    return int(match.group(1)), match.group(2).strip()


def _component_match(line: str) -> Optional[Component]:
    # First eight fields are numeric.  The remaining fields are comma-separated
    # word 9/10/11/12/13-14 fields.
    match = re.match(r"^\s*(-?\d+(?:\s+-?\d+){7})\s+(.*)$", line)
    if not match or "," not in line:
        return None
    values = [int(v) for v in match.group(1).split()]
    parts = match.group(2).split(",")
    while len(parts) < 5:
        parts.append("")
    last = parts[4].split()
    return Component(
        line_no=0,
        type_id=values[0],
        start_raw=(values[1], values[2], values[3]),
        end_raw=(values[4], values[5], values[6]),
        bore=values[7],
        item_index=_safe_int(parts[0]),
        weight=parts[1].strip(),
        word11=parts[2].strip(),
        skey=parts[3].strip(),
        word13=last[0] if last else "",
        word14=last[1] if len(last) > 1 else "",
        raw=line.rstrip("\r\n"),
    )


def _find_component_start(lines: list[str]) -> int:
    for index, line in enumerate(lines[10:], start=10):
        if _component_match(line):
            return index
    return len(lines)


def parse_header(lines: list[str], end_index: int) -> dict[int, str]:
    result: dict[int, str] = {}
    pending: Optional[int] = None
    for line in lines[10:end_index]:
        neg = _negative_match(line.strip())
        if not neg:
            continue
        record, text = neg
        if record == -1 and pending is not None:
            current = result.get(pending, "")
            # Most header values are identifiers split at a fixed width.  Join
            # those continuations directly.  The generator/version string is
            # the exception: an uppercase continuation starts a new word.
            if pending == -5 and current and text and text[0].isupper():
                current += " "
            result[pending] = normalize_text(f"{current}{text}")
        elif record != -1:
            result[record] = text
            pending = record
    return result


def parse_materials(lines: list[str], start_index: int) -> list[MaterialDef]:
    materials: list[MaterialDef] = []
    current: Optional[MaterialDef] = None
    mode = ""
    for line in lines[start_index:]:
        stripped = line.strip()
        if stripped == "999" or stripped.startswith("999 "):
            break
        # Keep any extra leading space in the payload.  Isogen wraps material
        # text at a fixed width, so that space tells us whether the wrap split
        # a word or occurred at a real word boundary.
        match = re.match(r"^\s*-20(?:\s)?(.*)$", line)
        if match:
            if current:
                current.code = normalize_text(current.code)
                current.description = normalize_text(current.description)
                materials.append(current)
            current = MaterialDef(len(materials) + 1, match.group(1).strip(), "")
            mode = "code"
            continue
        match = re.match(r"^\s*-21(?:\s)?(.*)$", line)
        if match and current:
            current.description = match.group(1)
            mode = "description"
            continue
        match = re.match(r"^\s*-1(?:\s)?(.*)$", line)
        if match and current:
            payload = match.group(1)
            if mode == "code":
                current.code = f"{current.code}{payload.strip()}"
            elif mode == "description":
                current.description = f"{current.description}{payload}"
    if current:
        current.code = normalize_text(current.code)
        current.description = normalize_text(current.description)
        materials.append(current)
    return materials


def parse_components(
    lines: list[str], start_index: int, material_index: int, initial_spec: str
) -> list[Component]:
    components: list[Component] = []
    current: Optional[Component] = None
    pending_negative: Optional[int] = None
    active_spec = initial_spec
    for index in range(start_index, material_index):
        line = lines[index]
        component = _component_match(line)
        if component:
            component.line_no = index + 1
            component.spec = active_spec
            components.append(component)
            current = component
            pending_negative = None
            continue

        neg = _negative_match(line.strip())
        if not neg:
            continue
        record, text = neg
        if record == -1:
            if current is not None and pending_negative is not None:
                _append_component_record(current, pending_negative, text)
            continue
        pending_negative = record
        if current is None:
            continue
        _append_component_record(current, record, text)
        if record == -26:
            active_spec = text
            current.spec_change_after = text
    return components


def _append_component_record(component: Component, record: int, text: str) -> None:
    if record == -39:
        component.unique_id = normalize_text(f"{component.unique_id}{text}")
    elif record == -22:
        component.tag = normalize_text(f"{component.tag}{text}")
    elif record == -30:
        component.pipeline_connection = normalize_text(f"{component.pipeline_connection}{text}")
    elif record == -31:
        component.equipment_connection = normalize_text(f"{component.equipment_connection}{text}")
    elif record == -37:
        component.user_message = normalize_text(f"{component.user_message}{text}")
    elif record == -26:
        component.spec_change_after = text
    elif record not in {-1}:
        component.notes.append(f"{record}: {text}")


def extract_idf(path: str | Path) -> dict[str, Any]:
    lines, encoding = read_idf(path)
    first_component = _find_component_start(lines)
    material_start = next(
        (i for i, line in enumerate(lines) if re.match(r"^\s*-20\s", line)),
        len(lines),
    )
    header = parse_header(lines, first_component)
    materials = parse_materials(lines, material_start)
    components = parse_components(
        lines,
        first_component,
        material_start,
        initial_spec=header.get(-11, ""),
    )
    material_by_index = {m.index: m for m in materials}
    detail_rows: list[dict[str, Any]] = []
    for component in components:
        if component.type_id in NON_PHYSICAL_TYPES or bom_control_value(component) == 1:
            continue
        material = material_by_index.get(component.item_index)
        description = material.description if material else ""
        nominal_diameter, wall_thickness, pressure_rating = parse_spec_dimensions(description, component.bore)
        sx, sy, sz = component.start_mm
        ex, ey, ez = component.end_mm
        detail_rows.append(
            {
                "来源文件名": Path(path).name,
                "公称管径": nominal_diameter,
                "壁厚等级": wall_thickness,
                "压力等级": pressure_rating,
                "材料指针": material.index if material else "",
                "材料编码": material.code if material else "",
                "材料描述": material.description if material else "",
                "数量": 1,
                "单位": "EA",
                "组件类型": component.type_name,
                "类型号": component.type_id,
                "管径/尺寸字段": component.bore or "",
                "SKEY": component.skey,
                "角度/螺栓长度": component.angle_or_bolt_length,
                "管道规格": component.spec,
                "唯一组件号": component.unique_id,
                "组件标签": component.tag,
                "起点E/W(mm)": sx,
                "起点N/S(mm)": sy,
                "起点标高(mm)": sz,
                "终点E/W(mm)": ex,
                "终点N/S(mm)": ey,
                "终点标高(mm)": ez,
                "几何长度(mm)": component.length_mm or "",
                "制造/安装分类": component.item_class,
                "流向": component.flow_direction,
                "主方向": component.primary_direction,
                "保温/伴热": component.insulation_tracing,
                "BOM显示": component.bom_status,
                "跨管线连接": component.pipeline_connection,
                "用户备注": component.user_message,
                "负记录备注": " | ".join(component.notes),
                "设备连接": header.get(-31, ""),
                "来源文件": Path(path).name,
                "原始IDF记录": component.raw,
                "IDF行号": component.line_no,
            }
        )

    grouped: dict[tuple[Any, ...], dict[str, Any]] = {}
    for row in (row for row in detail_rows if row["材料指针"] != ""):
        key = (row["材料指针"], row["公称管径"], row["壁厚等级"], row["压力等级"], row["管道规格"])
        if key not in grouped:
            grouped[key] = {
                "材料指针": row["材料指针"],
                "材料编码": row["材料编码"],
                "材料描述": row["材料描述"],
                "数量": 0,
                "单位": "EA",
                "组件类型": set(),
                "管径/尺寸字段": set(),
                "管道规格": set(),
                "唯一组件号": [],
                "来源文件": Path(path).name,
                "公称管径": row["公称管径"],
                "壁厚等级": row["壁厚等级"],
                "压力等级": row["压力等级"],
            }
        item = grouped[key]
        item["数量"] += 1
        if row["组件类型"]:
            item["组件类型"].add(row["组件类型"])
        if row["管径/尺寸字段"] != "":
            item["管径/尺寸字段"].add(str(row["管径/尺寸字段"]))
        if row["管道规格"]:
            item["管道规格"].add(row["管道规格"])
        if row["唯一组件号"]:
            item["唯一组件号"].append(row["唯一组件号"])

    summary_rows = []
    for item in grouped.values():
        summary_rows.append(
            {
                "来源文件名": item["来源文件"],
                "公称管径": item["公称管径"],
                "壁厚等级": item["壁厚等级"],
                "压力等级": item["压力等级"],
                "材料指针": item["材料指针"],
                "材料编码": item["材料编码"],
                "材料描述": item["材料描述"],
                "数量": item["数量"],
                "单位": item["单位"],
                "组件类型": "、".join(sorted(item["组件类型"])),
                "管径/尺寸字段": "、".join(sorted(item["管径/尺寸字段"], key=lambda x: float(x))),
                "管道规格": "、".join(sorted(item["管道规格"])),
                "唯一组件号": "、".join(item["唯一组件号"]),
                "来源文件": item["来源文件"],
            }
        )

    return {
        "path": str(Path(path).resolve()),
        "file_name": Path(path).name,
        "encoding": encoding,
        "header": header,
        "materials": materials,
        "components": components,
        "summary_rows": summary_rows,
        "detail_rows": detail_rows,
    }


def extract_idfs(paths: Iterable[str | Path]) -> dict[str, Any]:
    """批量读取多个 IDF，并保留每条记录的来源文件名。"""
    source_paths = [Path(path) for path in paths]
    if not source_paths:
        raise ValueError("至少需要选择一个 IDF 文件。")
    datasets = [extract_idf(path) for path in source_paths]
    detail_rows: list[dict[str, Any]] = []
    summary_rows: list[dict[str, Any]] = []
    for data in datasets:
        source_name = data["file_name"]
        for row in data["detail_rows"]:
            item = dict(row)
            item["来源文件名"] = source_name
            detail_rows.append(item)
        for row in data["summary_rows"]:
            item = dict(row)
            item["来源文件名"] = source_name
            summary_rows.append(item)
    return {
        "path": "；".join(str(path.resolve()) for path in source_paths),
        "file_name": "、".join(path.name for path in source_paths),
        "encoding": "；".join(f"{data['file_name']}:{data['encoding']}" for data in datasets),
        "header": {},
        "materials": [material for data in datasets for material in data["materials"]],
        "components": [component for data in datasets for component in data["components"]],
        "summary_rows": summary_rows,
        "detail_rows": detail_rows,
        "file_count": len(datasets),
    }


SUMMARY_HEADERS = [
    "来源文件名",
    "公称管径",
    "壁厚等级",
    "压力等级",
    "材料指针",
    "材料编码",
    "材料描述",
    "数量",
    "单位",
    "组件类型",
    "管径/尺寸字段",
    "管道规格",
    "唯一组件号",
    "来源文件",
]

DETAIL_HEADERS = [
    "来源文件名",
    "公称管径",
    "壁厚等级",
    "压力等级",
    "材料指针",
    "材料编码",
    "材料描述",
    "数量",
    "单位",
    "组件类型",
    "类型号",
    "管径/尺寸字段",
    "SKEY",
    "角度/螺栓长度",
    "管道规格",
    "唯一组件号",
    "组件标签",
    "起点E/W(mm)",
    "起点N/S(mm)",
    "起点标高(mm)",
    "终点E/W(mm)",
    "终点N/S(mm)",
    "终点标高(mm)",
    "几何长度(mm)",
    "制造/安装分类",
    "流向",
    "主方向",
    "保温/伴热",
    "BOM显示",
    "跨管线连接",
    "用户备注",
    "负记录备注",
    "设备连接",
    "来源文件",
    "原始IDF记录",
    "IDF行号",
]


HEADER_ALIASES = {
    "来源文件名": {"来源文件名", "sourcefilename", "filename", "文件名"},
    "公称管径": {"公称管径", "nominaldiameter", "dn", "diameter"},
    "壁厚等级": {"壁厚等级", "壁厚", "schedule", "sch", "wallthickness"},
    "压力等级": {"压力等级", "压力级别", "pressureclass", "pressure", "cl", "pn"},
    "材料指针": {"材料指针", "materialpointer", "materialindex", "itempointer", "index"},
    "材料编码": {"材料编码", "材料号", "itemcode", "materialcode", "partnumber", "item"},
    "材料描述": {"材料描述", "描述", "description", "materialdescription", "itemdescription", "title"},
    "数量": {"数量", "qty", "quantity", "count", "number"},
    "单位": {"单位", "unit", "uom"},
    "组件类型": {"组件类型", "类型", "componenttype", "type"},
    "管径/尺寸字段": {"公称管径mm", "公称管径", "管径", "管径尺寸", "size", "nominalsize", "bore", "diameter"},
    "管道规格": {"管道规格", "规格", "spec", "pipingspec", "piping specification"},
    "唯一组件号": {"唯一组件号", "组件号", "uniqueid", "uniquecomponentid", "pieceid", "componentid"},
    "来源文件": {"来源文件", "文件", "sourcefile", "source"},
}

DETAIL_FIELD_ALIASES = {
    "类型号": {"类型号", "typeid", "componenttypeid"},
    "SKEY": {"skey", "symbolkey", "符号键"},
    "角度/螺栓长度": {"角度螺栓长度", "angle", "boltlength", "angleorboltlength"},
    "组件标签": {"组件标签", "标签", "tag", "name"},
    "起点E/W(mm)": {"起点ewmm", "starteastwest", "startew"},
    "起点N/S(mm)": {"起点nsmm", "startnorthsouth", "startns"},
    "起点标高(mm)": {"起点标高mm", "startelevation", "startz"},
    "终点E/W(mm)": {"终点ewmm", "endeastwest", "endew"},
    "终点N/S(mm)": {"终点nsmm", "endnorthsouth", "endns"},
    "终点标高(mm)": {"终点标高mm", "endelevation", "endz"},
    "几何长度(mm)": {"几何长度mm", "lengthmm", "geometriclength"},
    "制造/安装分类": {"制造安装分类", "classification", "faberect"},
    "流向": {"流向", "flow", "flowdirection"},
    "主方向": {"主方向", "direction", "primarydirection"},
    "保温/伴热": {"保温伴热", "insulation", "tracing", "insulationtracing"},
    "BOM显示": {"bom显示", "bom", "bomstatus"},
    "跨管线连接": {"跨管线连接", "pipelineconnection", "connectionpipeline"},
    "用户备注": {"用户备注", "备注", "usermessage", "message"},
    "负记录备注": {"负记录备注", "negative record notes", "notes", "recordnotes"},
    "设备连接": {"设备连接", "equipmentconnection", "equipment"},
    "原始IDF记录": {"原始idf记录", "原始记录", "rawrecord", "rawidf"},
    "IDF行号": {"idf行号", "行号", "linenumber", "source line"},
}
for _field_name in DETAIL_HEADERS:
    HEADER_ALIASES.setdefault(_field_name, set()).update(DETAIL_FIELD_ALIASES.get(_field_name, {_field_name}))


def canonical_header(value: Any) -> str:
    text = normalize_text(str(value) if value is not None else "").lower()
    return re.sub(r"[\s_\-()/\\.]+", "", text)


def detect_header_row(ws, max_scan: int = 50) -> int:
    aliases = {canonical_header(alias) for values in HEADER_ALIASES.values() for alias in values}
    best_row, best_score = 1, -1
    for row in range(1, min(ws.max_row, max_scan) + 1):
        score = sum(1 for cell in ws[row] if canonical_header(cell.value) in aliases)
        if score > best_score:
            best_row, best_score = row, score
    return best_row


def _header_columns(ws, header_row: int) -> dict[str, int]:
    lookup: dict[str, int] = {}
    alias_to_field = {
        canonical_header(alias): field_name
        for field_name, aliases in HEADER_ALIASES.items()
        for alias in aliases
    }
    for col in range(1, ws.max_column + 1):
        field_name = alias_to_field.get(canonical_header(ws.cell(header_row, col).value))
        if field_name and field_name not in lookup:
            lookup[field_name] = col
    return lookup


def _copy_row_style(ws, source_row: int, target_row: int) -> None:
    if source_row <= 0 or source_row > ws.max_row:
        return
    for col in range(1, ws.max_column + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.font:
            target.font = copy(source.font)
        if source.fill:
            target.fill = copy(source.fill)
        if source.border:
            target.border = copy(source.border)


def _last_data_row(ws, header_row: int) -> int:
    last = header_row
    for row in range(header_row + 1, ws.max_row + 1):
        if any(ws.cell(row, col).value not in (None, "") for col in range(1, ws.max_column + 1)):
            last = row
    return last


def write_custom_sheet(ws, rows: list[dict[str, Any]], headers: list[str], header_row: int) -> None:
    columns = _header_columns(ws, header_row)
    missing = [field_name for field_name in headers if field_name not in columns]
    next_col = ws.max_column + 1
    for field_name in missing:
        ws.cell(header_row, next_col).value = field_name
        columns[field_name] = next_col
        next_col += 1

    start_row = _last_data_row(ws, header_row) + 1
    style_row = start_row - 1 if start_row - 1 > header_row else None
    for offset, row_data in enumerate(rows):
        row_no = start_row + offset
        if style_row:
            _copy_row_style(ws, style_row, row_no)
        for field_name in headers:
            ws.cell(row_no, columns[field_name]).value = row_data.get(field_name, "")
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max(ws.max_column, next_col - 1))}{max(start_row + len(rows) - 1, header_row)}"


def _apply_table_style(ws, ref: str, name: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(table)


def _write_matrix(ws, start_row: int, headers: list[str], rows: Iterable[dict[str, Any]]) -> int:
    for col, header in enumerate(headers, 1):
        cell = ws.cell(start_row, col, header)
        cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    row_no = start_row + 1
    for row in rows:
        for col, header in enumerate(headers, 1):
            ws.cell(row_no, col, row.get(header, ""))
        row_no += 1
    return row_no - 1


def _format_sheet(ws, header_row: int, end_row: int) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ws.max_column)}{max(end_row, header_row)}"
    thin = Side(style="thin", color="D9E2F3")
    for row in ws.iter_rows(min_row=header_row, max_row=max(end_row, header_row), min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            if cell.row > header_row:
                cell.border = Border(bottom=thin)
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        width = 12
        for row in range(header_row, min(end_row, header_row + 80) + 1):
            value = ws.cell(row, col).value
            if value is not None:
                width = min(42, max(width, len(str(value)) + 2))
        ws.column_dimensions[letter].width = width


def export_new_workbook(data: dict[str, Any], output_path: str | Path) -> None:
    wb = Workbook()
    detail = wb.active
    detail.title = "管件明细"
    summary = wb.create_sheet("材料汇总")
    info = wb.create_sheet("文件信息")

    detail["A1"] = "IDF管件明细（每个组件一行）"
    detail["A1"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E78")
    detail_end = _write_matrix(detail, 3, DETAIL_HEADERS, data["detail_rows"])
    _format_sheet(detail, 3, detail_end)
    if detail_end >= 4:
        _apply_table_style(detail, f"A3:{get_column_letter(len(DETAIL_HEADERS))}{detail_end}", "FittingDetail")

    summary["A1"] = "IDF材料汇总（按材料编码合并）"
    summary["A1"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E78")
    summary["A2"] = f"来源文件：{data['file_name']}"
    summary["A2"].font = Font(name="Microsoft YaHei", size=10, italic=True, color="666666")
    summary_end = _write_matrix(summary, 4, SUMMARY_HEADERS, data["summary_rows"])
    _format_sheet(summary, 4, summary_end)
    if summary_end >= 5:
        _apply_table_style(summary, f"A4:{get_column_letter(len(SUMMARY_HEADERS))}{summary_end}", "MaterialSummary")

    info["A1"] = "文件信息"
    info["A1"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E78")
    info_rows = [
        ("来源文件", data["file_name"]),
        ("完整路径", data["path"]),
        ("读取编码", data["encoding"]),
        ("材料定义数量", len(data["materials"])),
        ("管件明细行数", len(data["detail_rows"])),
        ("材料汇总行数", len(data["summary_rows"])),
    ]
    for row, (label, value) in enumerate(info_rows, 3):
        info.cell(row, 1, label).font = Font(name="Microsoft YaHei", bold=True)
        info.cell(row, 2, value)
    row = len(info_rows) + 5
    info.cell(row, 1, "IDF负记录").font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF")
    info.cell(row, 1).fill = PatternFill("solid", fgColor="1F4E78")
    info.cell(row, 2).value = "规则含义"
    info.cell(row, 2).font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF")
    info.cell(row, 2).fill = PatternFill("solid", fgColor="1F4E78")
    for offset, (record, label) in enumerate(KNOWN_HEADER_RECORDS.items(), 1):
        info.cell(row + offset, 1, record)
        info.cell(row + offset, 2, f"{label}: {data['header'].get(record, '')}")
    _format_sheet(info, row, row + len(KNOWN_HEADER_RECORDS))

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def export_to_custom_template(
    data: dict[str, Any], template_path: str | Path, output_path: str | Path,
    sheet_name: str = "", header_row: int = 0,
) -> None:
    template_path = Path(template_path)
    output_path = Path(output_path)
    if template_path.resolve() == output_path.resolve():
        raise ValueError("输出文件不能与自定义模板相同，请另选输出路径。")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    wb = load_workbook(output_path, keep_vba=output_path.suffix.lower() == ".xlsm")
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb[wb.sheetnames[0]]
    header_row = header_row or detect_header_row(ws)
    write_custom_sheet(ws, data["detail_rows"], DETAIL_HEADERS, header_row)

    # If the template already has a detail sheet, populate it too without
    # changing the template's other worksheets.
    for detail_sheet_name in ("管件明细", "组件明细"):
        if detail_sheet_name not in wb.sheetnames or wb[detail_sheet_name] is ws:
            continue
        detail_ws = wb[detail_sheet_name]
        detail_header = detect_header_row(detail_ws)
        detail_columns = _header_columns(detail_ws, detail_header)
        for field_name in DETAIL_HEADERS:
            if field_name not in detail_columns:
                col = detail_ws.max_column + 1
                detail_ws.cell(detail_header, col, field_name)
                detail_columns[field_name] = col
        start = _last_data_row(detail_ws, detail_header) + 1
        for offset, row_data in enumerate(data["detail_rows"]):
            row_no = start + offset
            for field_name in DETAIL_HEADERS:
                detail_ws.cell(row_no, detail_columns[field_name], row_data.get(field_name, ""))
        detail_ws.freeze_panes = f"A{detail_header + 1}"
    wb.save(output_path)


def _default_output(idf_path: str | Path) -> str:
    path = Path(idf_path)
    return str(path.with_name(f"{path.stem}_材料清单.xlsx"))


class MaterialExtractorApp:
    def __init__(self, root):
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk

        self.tk = tk
        self.filedialog = filedialog
        self.messagebox = messagebox
        self.ttk = ttk
        self.root = root
        self.root.title("IDF材料提取工具")
        self.root.geometry("980x650")
        self.data: Optional[dict[str, Any]] = None
        self.template_wb = None

        self.idf_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="new")
        self.sheet_var = tk.StringVar()
        self.header_row_var = tk.StringVar(value="自动识别")
        self.status_var = tk.StringVar(value="请选择一个 IDF 文件。")

        root.columnconfigure(1, weight=1)
        root.rowconfigure(5, weight=1)
        self._build_ui()

    def _build_ui(self):
        tk, ttk = self.tk, self.ttk
        pad = {"padx": 8, "pady": 6}
        ttk.Label(self.root, text="IDF文件").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(self.root, textvariable=self.idf_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(self.root, text="选择…", command=self.choose_idf).grid(row=0, column=2, **pad)

        ttk.Label(self.root, text="输出方式").grid(row=1, column=0, sticky="w", **pad)
        mode_frame = ttk.Frame(self.root)
        mode_frame.grid(row=1, column=1, columnspan=2, sticky="w", **pad)
        ttk.Radiobutton(mode_frame, text="新建Excel", variable=self.mode_var, value="new", command=self.refresh_mode).pack(side="left", padx=(0, 20))
        ttk.Radiobutton(mode_frame, text="写入自定义模板", variable=self.mode_var, value="custom", command=self.refresh_mode).pack(side="left")

        ttk.Label(self.root, text="自定义模板").grid(row=2, column=0, sticky="w", **pad)
        self.template_entry = ttk.Entry(self.root, textvariable=self.template_var, state="disabled")
        self.template_entry.grid(row=2, column=1, sticky="ew", **pad)
        self.template_btn = ttk.Button(self.root, text="选择…", command=self.choose_template, state="disabled")
        self.template_btn.grid(row=2, column=2, **pad)

        ttk.Label(self.root, text="模板工作表").grid(row=3, column=0, sticky="w", **pad)
        self.sheet_combo = ttk.Combobox(self.root, textvariable=self.sheet_var, state="disabled")
        self.sheet_combo.grid(row=3, column=1, sticky="ew", **pad)
        ttk.Label(self.root, text="表头行（可填数字）").grid(row=3, column=2, sticky="w", **pad)
        self.header_entry = ttk.Entry(self.root, textvariable=self.header_row_var, width=12, state="disabled")
        self.header_entry.grid(row=3, column=3, sticky="w", padx=8, pady=6)

        ttk.Label(self.root, text="输出文件").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(self.root, textvariable=self.output_var).grid(row=4, column=1, sticky="ew", **pad)
        ttk.Button(self.root, text="选择…", command=self.choose_output).grid(row=4, column=2, **pad)

        preview_frame = ttk.LabelFrame(self.root, text="识别预览")
        preview_frame.grid(row=5, column=0, columnspan=4, sticky="nsew", padx=8, pady=8)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.preview = ttk.Treeview(preview_frame, columns=("code", "desc", "type", "dn", "sch", "pressure", "bores", "spec", "uid"), show="headings")
        for col, title, width in (("code", "材料编码", 180), ("desc", "材料描述", 330), ("type", "组件类型", 100), ("dn", "公称管径", 90), ("sch", "壁厚等级", 90), ("pressure", "压力等级", 90), ("bores", "管径/尺寸", 100), ("spec", "管道规格", 150), ("uid", "唯一组件号", 130)):
            self.preview.heading(col, text=title)
            self.preview.column(col, width=width, anchor="w")
        self.preview.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.preview.configure(yscrollcommand=scroll.set)

        bottom = ttk.Frame(self.root)
        bottom.grid(row=6, column=0, columnspan=4, sticky="ew", padx=8, pady=8)
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        ttk.Button(bottom, text="提取并输出", command=self.run_export).grid(row=0, column=1, padx=8)

    def refresh_mode(self):
        state = "normal" if self.mode_var.get() == "custom" else "disabled"
        self.template_entry.configure(state=state)
        self.template_btn.configure(state=state)
        self.sheet_combo.configure(state="readonly" if self.mode_var.get() == "custom" else "disabled")
        self.header_entry.configure(state=state)

    def choose_idf(self):
        path = self.filedialog.askopenfilename(filetypes=[("IDF files", "*.idf"), ("All files", "*.*")])
        if not path:
            return
        self.idf_var.set(path)
        self.output_var.set(_default_output(path))
        try:
            self.data = extract_idf(path)
            self.fill_preview()
            self.status_var.set(f"已识别 {len(self.data['summary_rows'])} 种材料，{len(self.data['detail_rows'])} 条组件明细。")
        except Exception as exc:
            self.data = None
            self.messagebox.showerror("读取失败", str(exc))

    def choose_template(self):
        path = self.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")])
        if not path:
            return
        try:
            wb = load_workbook(path, read_only=True, keep_vba=path.lower().endswith(".xlsm"))
            self.template_wb = wb
            self.template_var.set(path)
            self.sheet_combo["values"] = wb.sheetnames
            if wb.sheetnames:
                self.sheet_var.set(wb.sheetnames[0])
        except Exception as exc:
            self.messagebox.showerror("模板读取失败", str(exc))

    def choose_output(self):
        path = self.filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("Macro-enabled Excel", "*.xlsm")],
        )
        if path:
            self.output_var.set(path)

    def fill_preview(self):
        for item in self.preview.get_children():
            self.preview.delete(item)
        if not self.data:
            return
        for row in self.data["detail_rows"]:
            self.preview.insert("", "end", values=(row["材料编码"], row["材料描述"], row["组件类型"], row["公称管径"], row["壁厚等级"], row["压力等级"], row["管径/尺寸字段"], row["管道规格"], row["唯一组件号"]))

    def run_export(self):
        if not self.data:
            self.messagebox.showwarning("缺少文件", "请先选择并读取 IDF 文件。")
            return
        output = self.output_var.get().strip() or _default_output(self.idf_var.get())
        try:
            if self.mode_var.get() == "new":
                export_new_workbook(self.data, output)
            else:
                template = self.template_var.get().strip()
                if not template:
                    raise ValueError("请选择自定义模板。")
                header_text = self.header_row_var.get().strip()
                header_row = int(header_text) if header_text.isdigit() else 0
                export_to_custom_template(self.data, template, output, self.sheet_var.get(), header_row)
            self.output_var.set(output)
            self.status_var.set(f"已输出：{output}")
            self.messagebox.showinfo("完成", f"材料信息已输出到：\n{output}")
        except Exception as exc:
            self.messagebox.showerror("输出失败", str(exc))


def run_cli(args: argparse.Namespace) -> None:
    data = extract_idf(args.idf)
    if args.mode == "new":
        export_new_workbook(data, args.output or _default_output(args.idf))
    else:
        if not args.template:
            raise ValueError("custom 模式必须提供 --template")
        header_row = args.header_row or 0
        export_to_custom_template(data, args.template, args.output, args.sheet_name or "", header_row)
    print(f"已输出 {len(data['summary_rows'])} 种材料、{len(data['detail_rows'])} 条组件明细。")


def main() -> None:
    parser = argparse.ArgumentParser(description="从 Isogen/AVEVA IDF 文件提取材料并输出 Excel")
    parser.add_argument("--idf", help="IDF文件路径")
    parser.add_argument("--mode", choices=("new", "custom"), default="new", help="输出模式")
    parser.add_argument("--template", help="自定义Excel模板路径")
    parser.add_argument("--sheet-name", default="", help="自定义模板目标工作表")
    parser.add_argument("--header-row", type=int, default=0, help="模板表头行号，0=自动识别")
    parser.add_argument("--output", help="输出Excel路径")
    args = parser.parse_args()
    if args.idf:
        run_cli(args)
        return
    try:
        import tkinter as tk
    except ImportError as exc:
        raise SystemExit("当前Python没有Tkinter，请安装带Tk支持的Python，或使用命令行参数运行。") from exc
    root = tk.Tk()
    MaterialExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
