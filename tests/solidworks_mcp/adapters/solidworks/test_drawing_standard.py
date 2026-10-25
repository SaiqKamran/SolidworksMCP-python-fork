"""Tests for the A3 drawing standard (``adapters/solidworks/drawing_standard.py``).

The sheet maths is tested directly. The mixin runs against a fake SolidWorks:
``Create3rdAngleViews2`` adds front, top and right views with real outlines,
``CreateDrawViewFromModelView3`` adds an isometric, notes report a text box
from their position, and the adapter methods the generator calls (open,
material, properties, save, export, dimensions) are recorded stubs.
"""

from __future__ import annotations

import builtins
import os
from types import SimpleNamespace
from typing import Any

import pytest

from solidworks_mcp.adapters.base import AdapterResult, AdapterResultStatus
from solidworks_mcp.adapters.solidworks import dimension_layout as lay
from solidworks_mcp.adapters.solidworks import drawing_standard as std
from solidworks_mcp.adapters.solidworks.drawing_standard import (
    SolidWorksDrawingStandardMixin,
)


def _ok(data: Any = None) -> AdapterResult[Any]:
    return AdapterResult(status=AdapterResultStatus.SUCCESS, data=data)


def _err(error: str) -> AdapterResult[Any]:
    return AdapterResult(status=AdapterResultStatus.ERROR, error=error)


# ---- sheet maths ----------------------------------------------------------------


class TestSheetMaths:
    PLATE = {"front": (0.12, 0.01), "plan": (0.12, 0.07), "side": (0.07, 0.01)}

    def test_plate_fills_the_sheet(self) -> None:
        scale, cells = std.choose_layout(self.PLATE)
        assert scale == (1.5, 1)
        for name in ("front", "side", "plan", "iso"):
            assert lay.inside(cells[name], std.VIEW_ZONE)
            assert not lay.box_overlap(cells[name], std.TITLE_BLOCK)

    def test_cells_do_not_overlap(self) -> None:
        _, cells = std.choose_layout(self.PLATE)
        names = ["front", "side", "plan", "iso", "notes"]
        for i, a in enumerate(names):
            for b in names[i + 1 :]:
                assert not lay.box_overlap(cells[a], cells[b], -1e-9), (a, b)

    def test_large_part_uses_a_smaller_standard_scale(self) -> None:
        scale, _ = std.choose_layout(
            {"front": (0.4, 0.3), "plan": (0.4, 0.25), "side": (0.25, 0.3)}
        )
        assert scale in std.STANDARD_SCALES and scale[0] / scale[1] < 1

    def test_depth_comes_from_the_plan_without_a_side_view(self) -> None:
        with_side = std.plan_cells(self.PLATE, (1, 1))
        no_side = std.plan_cells({"front": (0.12, 0.01), "plan": (0.12, 0.07)}, (1, 1))
        assert with_side == no_side

    def test_no_front_view_or_no_fit_gives_none(self) -> None:
        assert std.plan_cells({"plan": (0.1, 0.1)}, (1, 1)) is None
        assert std.plan_cells({"front": (1.0, 1.0)}, (1, 1)) is None

    def test_a_part_too_big_for_any_scale_gets_the_smallest_and_default_cells(
        self,
    ) -> None:
        scale, cells = std.choose_layout({"front": (100.0, 100.0)})
        assert scale == std.STANDARD_SCALES[-1]
        assert cells == std.CELLS

    def test_iso_fit_and_labels(self) -> None:
        assert std.scale_label((1, 2)) == "1:2"
        assert std.scale_label((2.5, 1)) == "2.5:1"
        _, cells = std.choose_layout(self.PLATE)
        assert std.iso_fit_scale((0.15, 0.09), cells["iso"]) in std.STANDARD_SCALES
        assert std.iso_fit_scale((100.0, 100.0), cells["iso"]) == (1, 100)
        assert std.content_box((0, 0, 1, 1)) == (0, std.CAPTION_STRIP, 1, 1)

    def test_title_block_properties(self) -> None:
        model, drawing = std.title_block_properties(
            model_file=r"C:\parts\Mounting Plate.SLDPRT",
            title="Mounting Plate - Rev A",
            subtitle="Part drawing",
            author="saiq",
            drawn_date="17 Oct 2026",
            revision="A",
            company="",
        )
        assert model == {
            "Description": "Mounting Plate - Rev A",
            "Material": '"SW-Material@Mounting Plate.SLDPRT"',
            "Weight": '"SW-Mass@Mounting Plate.SLDPRT"',
        }
        assert drawing == {
            "DrawnBy": "saiq",
            "DrawnDate": "17 Oct 2026",
            "Revision": "A",
            "COMPANYNAME": "Part drawing",
        }

    def test_empty_drawing_properties_are_left_out(self) -> None:
        _, drawing = std.title_block_properties(
            model_file="a.SLDPRT",
            title="A",
            author="",
            drawn_date="1 Jan 2026",
            revision="A",
            company="",
        )
        assert drawing == {"DrawnDate": "1 Jan 2026", "Revision": "A"}


class TestDoubleArray:
    def test_variant_on_windows(self, monkeypatch) -> None:
        fake_w32 = SimpleNamespace(VARIANT=lambda vt, values: ("VARIANT", vt, values))
        monkeypatch.setitem(
            __import__("sys").modules,
            "pythoncom",
            SimpleNamespace(VT_ARRAY=0x2000, VT_R8=5),
        )
        monkeypatch.setitem(
            __import__("sys").modules, "win32com", SimpleNamespace(client=fake_w32)
        )
        monkeypatch.setitem(__import__("sys").modules, "win32com.client", fake_w32)
        assert std._double_array([1, 2]) == ("VARIANT", 0x2005, [1.0, 2.0])

    def test_plain_tuple_without_pywin32(self, monkeypatch) -> None:
        real_import = builtins.__import__

        def no_pythoncom(name, *args, **kwargs):
            if name == "pythoncom":
                raise ImportError(name)
            return real_import(name, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", no_pythoncom)
        assert std._double_array([1.0, 2.0]) == (1.0, 2.0)


class TestFindSheetFormat:
    def test_configured_folder_first(self, tmp_path, monkeypatch) -> None:
        (tmp_path / "a3 - iso.slddrt").write_bytes(b"")
        monkeypatch.setattr(std.glob, "glob", lambda _pattern: [])
        adapter = _Harness()
        adapter.swApp.prefs[std._SW_PREF_SHEET_FORMAT_FOLDER] = f" ; {tmp_path}"
        assert std.find_sheet_format(adapter) == str(tmp_path / "a3 - iso.slddrt")

    def test_programdata_fallback_and_missing(self, tmp_path, monkeypatch) -> None:
        (tmp_path / "a3 - iso.slddrt").write_bytes(b"")
        monkeypatch.setattr(std.glob, "glob", lambda _pattern: [str(tmp_path)])
        no_app = SimpleNamespace(swApp=None)
        assert std.find_sheet_format(no_app) == str(tmp_path / "a3 - iso.slddrt")
        assert std.find_sheet_format(no_app, "missing.slddrt") is None


# ---- fake SolidWorks ------------------------------------------------------------


def _value(raw: Any) -> Any:
    """What a VARIANT (or the plain fallback) carries."""
    return getattr(raw, "value", raw)


class _View:
    def __init__(self, name: str, orientation: str, outline: tuple) -> None:
        self.name, self.orientation, self.outline = name, orientation, outline
        self.ScaleDecimal = 1.0
        self.UseSheetScale = True
        self.ReferencedConfiguration = "Default"
        self._position = ((outline[0] + outline[2]) / 2, (outline[1] + outline[3]) / 2)
        self.bom: tuple | None = None
        self.bom_result: Any = SimpleNamespace(RowCount=3)
        self.stuck = False

    @property
    def Position(self):  # noqa: N802
        return self._position

    @Position.setter
    def Position(self, raw):  # noqa: N802
        if not self.stuck:
            self._position = tuple(_value(raw))

    @property
    def ScaleRatio(self):  # noqa: N802
        return (1.0, 1.0)

    @ScaleRatio.setter
    def ScaleRatio(self, raw):  # noqa: N802
        self.ratio = tuple(_value(raw))

    def GetOutline(self):  # noqa: N802
        return self.outline

    def GetOrientationName(self):  # noqa: N802
        return self.orientation

    def InsertBomTable4(self, *args):  # noqa: N802
        self.bom = args
        return self.bom_result


class _Sheet:
    def __init__(self) -> None:
        self.scale: tuple | None = None
        self.accept = True

    def GetName(self):  # noqa: N802
        return "Sheet1"

    def SetScale(self, num, den, *_):  # noqa: N802
        if self.accept:
            self.scale = (num, den)
        return self.accept

    def GetProperties2(self):  # noqa: N802
        if self.scale is None:
            return None
        return (8, 12, self.scale[0], self.scale[1], False, 0.42, 0.297, True)


class _Ann:
    """A note's ``IAnnotation``: its box is 4 mm wide per 2 characters."""

    def __init__(self, text: str) -> None:
        self.text, self.pos = text, (0.0, 0.0, 0.0)
        self.format: Any = SimpleNamespace(CharHeight=0.0, Bold=False)
        self.no_position = False

    def GetTextFormat(self, _i):  # noqa: N802
        return self.format

    def SetTextFormat(self, *_):  # noqa: N802
        return True

    def SetPosition(self, x, y, z):  # noqa: N802
        self.pos = (x, y, z)
        return True

    def GetPosition(self):  # noqa: N802
        return None if self.no_position else self.pos


class _Drawing:
    """``IModelDoc2`` + ``IDrawingDoc`` of a new drawing."""

    def __init__(self) -> None:
        self.views: dict[str, _View] = {}
        self.sheet = _Sheet()
        self.setup_ok = True
        self.third_angle: str = "views"  # "views", "none", "false_but_views"
        self.iso = True
        self.notes: list[_Ann] = []
        self.note_mode = "ok"  # "ok", "no_note", "no_annotation", "no_format"
        self.prefs: dict[int, Any] = {}

    def GetType(self):  # noqa: N802
        return 3

    def GetCurrentSheet(self):  # noqa: N802
        return self.sheet

    def SetupSheet5(self, *_):  # noqa: N802
        return self.setup_ok

    def SetUserPreferenceIntegerValue(self, key, value):  # noqa: N802
        self.prefs[key] = value
        return True

    def Create3rdAngleViews2(self, _path):  # noqa: N802
        if self.third_angle == "none":
            return False
        self.views["Drawing View1"] = _View(
            "Drawing View1", "*Front", (0.05, 0.1, 0.17, 0.11)
        )
        self.views["Drawing View2"] = _View(
            "Drawing View2", "*Top", (0.05, 0.13, 0.17, 0.2)
        )
        self.views["Drawing View3"] = _View(
            "Drawing View3", "*Right", (0.2, 0.1, 0.27, 0.11)
        )
        return self.third_angle == "views"

    def CreateDrawViewFromModelView3(self, *_):  # noqa: N802
        if self.iso:
            self.views["Drawing View4"] = _View(
                "Drawing View4", "*Isometric", (0.3, 0.2, 0.45, 0.29)
            )
        return None

    def InsertNote(self, text):  # noqa: N802
        if self.note_mode == "no_note":
            return None
        ann = _Ann(text)
        if self.note_mode == "no_format":
            ann.format = None
        self.notes.append(ann)
        annotation = None if self.note_mode == "no_annotation" else ann
        return SimpleNamespace(GetAnnotation=lambda: annotation)

    def EditRebuild3(self):  # noqa: N802
        return True

    def ClearSelection2(self, _all):  # noqa: N802
        return None


class _App:
    def __init__(self) -> None:
        self.prefs: dict[int, Any] = {}
        self.toggles: dict[int, Any] = {std._SW_SHOW_SHEET_FORMAT_DIALOG: True}
        self.open_docs: dict[str, Any] = {}
        self.activated: list[str] = []

    def GetUserPreferenceStringValue(self, key):  # noqa: N802
        return self.prefs.get(key, "")

    def GetUserPreferenceToggle(self, key):  # noqa: N802
        return self.toggles.get(key)

    def SetUserPreferenceToggle(self, key, value):  # noqa: N802
        self.toggles[key] = value

    def GetOpenDocumentByName(self, path):  # noqa: N802
        return self.open_docs.get(path)

    def ActivateDoc3(self, title, *_):  # noqa: N802
        self.activated.append(title)


@pytest.fixture(autouse=True)
def _plain_com(monkeypatch) -> None:
    """Fake objects need no flagging or wrapping; views come from the fake."""
    monkeypatch.setattr(
        std, "_sw_type_info", SimpleNamespace(flagged=lambda obj, *_: obj)
    )
    monkeypatch.setattr(std, "_as_com", lambda _adapter, obj, _iface: obj)
    monkeypatch.setattr(std, "_view_names", lambda _a, drawing: list(drawing.views))
    monkeypatch.setattr(
        std,
        "_resolve_drawing_view",
        lambda _a, drawing, name: drawing.views.get(name),
    )


class _Harness(SolidWorksDrawingStandardMixin):
    """The mixin plus the PyWin32Adapter members and methods it calls."""

    def __init__(self, model: Any = None) -> None:
        self.currentModel = model
        self.swApp = _App()
        self.calls: list[tuple] = []
        self.drawing = _Drawing()
        self.results: dict[str, AdapterResult[Any]] = {}
        self.final_overlaps = 0

    # PyWin32Adapter plumbing
    def _handle_com_operation(self, name, operation, *args):
        try:
            data = operation(*args)
        except Exception as e:
            return _err(f"Error in {name}: {e}")
        return _ok(data)

    def _attempt(self, operation, default=None):
        try:
            return operation()
        except Exception:
            return default

    def _require_drawing(self):
        if getattr(self.currentModel, "GetType", lambda: 0)() != 3:
            return _err("This operation requires a drawing document")
        return None

    # the adapter methods the generator awaits
    def _result(self, name: str, *args: Any) -> AdapterResult[Any]:
        self.calls.append((name, *args))
        return self.results.get(name, _ok())

    async def open_model(self, path):
        self.currentModel = SimpleNamespace(path=path)
        return self._result("open_model", path)

    async def get_material(self):
        self.calls.append(("get_material",))
        return self.results.get("get_material", _ok({"material": "AISI 1020"}))

    async def set_custom_property(self, name, value):
        return self._result("set_custom_property", name, value)

    async def save_file(self, path=None):
        return self._result("save_file", path)

    async def create_drawing(self):
        self.calls.append(("create_drawing", dict(self.swApp.toggles)))
        result = self.results.get("create_drawing", _ok())
        if result.is_success:
            self.currentModel = self.drawing
        return result

    async def set_units(self, units):
        return self._result("set_units", units)

    async def export_file(self, path, fmt):
        return self._result("export_file", path, fmt)

    # the dimension methods from SolidWorksDocumentationMixin
    def _import_dimensions(self, view_name, include_hole_wizard, only_views=None):
        self.calls.append(("import_dimensions", tuple(only_views or ())))
        return {"dimensions_added": 6}

    def _arrange_layout(self, view_name, obstacles, work_area, move):
        self.calls.append(("arrange_layout", move, tuple(sorted(obstacles))))
        return {"overlaps_after": 0 if move else self.final_overlaps}

    def _collect_layout(self, drawing, names):
        return {}, [], {}

    def _dimension_geometry(self, an, key, view_name):
        if an.text == "BROKEN":
            raise RuntimeError("no display data")
        x, y = an.pos[0], an.pos[1]
        return SimpleNamespace(text_box=(x, y - 0.003, x + 0.002 * len(an.text), y))


# ---- BOM and sheet scale --------------------------------------------------------


def _drawing_harness() -> _Harness:
    adapter = _Harness()
    adapter.currentModel = adapter.drawing
    adapter.drawing.Create3rdAngleViews2("x")
    return adapter


class TestInsertBomTable:
    @pytest.mark.asyncio
    async def test_default_template_and_anchor(self, tmp_path) -> None:
        (tmp_path / "bom-standard.sldbomtbt").write_bytes(b"")
        adapter = _drawing_harness()
        adapter.swApp.prefs[std._SW_PREF_BOM_TEMPLATES] = f"{tmp_path};other"

        result = await adapter.insert_bom_table("Drawing View1")

        template = str(tmp_path / "bom-standard.sldbomtbt")
        assert result.data == {
            "view": "Drawing View1",
            "bom_type": "top_level_only",
            "template": template,
            "row_count": 3,
        }
        bom = adapter.drawing.views["Drawing View1"].bom
        assert bom == (True, 0.0, 0.0, 1, 2, "Default", template, False, 0, False)

    @pytest.mark.asyncio
    async def test_given_template_position_and_indented_numbering(self) -> None:
        adapter = _drawing_harness()

        result = await adapter.insert_bom_table(
            "Drawing View1", "C:/t.sldbomtbt", "indented", 100.0, 50.0
        )

        assert result.is_success
        bom = adapter.drawing.views["Drawing View1"].bom
        assert bom[:3] == (False, 0.1, 0.05)
        assert bom[4] == 3 and bom[6] == "C:/t.sldbomtbt" and bom[8] == 1

    @pytest.mark.asyncio
    async def test_a_part_view_says_a_bom_needs_an_assembly(self) -> None:
        adapter = _drawing_harness()
        adapter.drawing.views["Drawing View1"].bom_result = None

        no_template = await adapter.insert_bom_table("Drawing View1")
        with_template = await adapter.insert_bom_table(
            "Drawing View1", "C:/t.sldbomtbt"
        )

        assert "A BOM needs a view of an assembly and a BOM template (none found)." in (
            no_template.error
        )
        assert with_template.error.endswith("A BOM needs a view of an assembly.")

    @pytest.mark.asyncio
    async def test_errors(self) -> None:
        adapter = _drawing_harness()

        unknown = await adapter.insert_bom_table("Side")
        bad_type = await adapter.insert_bom_table("Drawing View1", bom_type="flat")
        adapter.currentModel = SimpleNamespace(GetType=lambda: 1)
        not_drawing = await adapter.insert_bom_table("Drawing View1")

        assert unknown.error == (
            "Error in insert_bom_table: No view named 'Side'. Views: "
            "Drawing View1, Drawing View2, Drawing View3"
        )
        assert bad_type.error.startswith("bom_type must be")
        assert not_drawing.error == "This operation requires a drawing document"

    @pytest.mark.asyncio
    async def test_no_views_at_all(self) -> None:
        adapter = _Harness()
        adapter.currentModel = adapter.drawing
        result = await adapter.insert_bom_table("Side")
        assert result.error.endswith("Views: none")


class TestSetSheetScale:
    @pytest.mark.asyncio
    async def test_sets_and_reports_the_scale(self) -> None:
        adapter = _drawing_harness()
        result = await adapter.set_sheet_scale(1, 2)
        assert result.data == {"sheet": "Sheet1", "scale": "1:2", "requested": "1:2"}
        assert adapter.drawing.sheet.scale == (1.0, 2.0)

    @pytest.mark.asyncio
    async def test_errors(self) -> None:
        adapter = _drawing_harness()
        adapter.drawing.sheet.accept = False

        refused = await adapter.set_sheet_scale(1, 2)
        adapter.drawing.sheet.accept = True
        adapter.drawing.sheet.GetProperties2 = lambda: None
        unreadable = await adapter.set_sheet_scale(1, 2)
        bad = await adapter.set_sheet_scale(0, 2)
        adapter.currentModel = None
        not_drawing = await adapter.set_sheet_scale(1, 2)

        assert (
            refused.error == "Error in set_sheet_scale: ISheet.SetScale returned False"
        )
        assert unreadable.error.endswith("could not be read back")
        assert bad.error == "Scale values must be positive"
        assert not_drawing.error == "This operation requires a drawing document"


# ---- small helpers --------------------------------------------------------------


class TestHelpers:
    def test_set_view_position_reads_back(self) -> None:
        adapter = _Harness()
        view = _View("v", "*Front", (0, 0, 0.1, 0.1))
        assert adapter._set_view_position(view, 0.2, 0.15)
        view.stuck = True
        assert not adapter._set_view_position(view, 0.3, 0.15)

    def test_use_open_document(self, tmp_path) -> None:
        adapter = _Harness()
        path = os.path.normcase(str(tmp_path / "a.SLDPRT"))
        assert not adapter._use_open_document(path)

        adapter.swApp.open_docs[path] = SimpleNamespace(
            GetPathName=lambda: str(tmp_path / "other.SLDPRT"), GetTitle=lambda: "a"
        )
        assert not adapter._use_open_document(path)

        doc = SimpleNamespace(GetPathName=lambda: path, GetTitle=lambda: "a.SLDPRT")
        adapter.swApp.open_docs[path] = doc
        assert adapter._use_open_document(path)
        assert adapter.currentModel is doc and adapter.swApp.activated == ["a.SLDPRT"]

        untitled = SimpleNamespace(GetPathName=lambda: path, GetTitle=lambda: "")
        adapter.swApp.open_docs[path] = untitled
        assert adapter._use_open_document(path)
        assert adapter.swApp.activated == ["a.SLDPRT"]

    def test_view_outline(self) -> None:
        adapter = _Harness()
        assert adapter._view_outline(_View("v", "", (0, 0, 1, 2))) == (0, 0, 1, 2)
        assert adapter._view_outline(SimpleNamespace(GetOutline=lambda: None)) is None

    def test_insert_sheet_note(self) -> None:
        adapter = _Harness(model=_Drawing())
        an = adapter._insert_sheet_note("NOTES", 0.02, 0.07, 0.003, True)
        assert an.pos == (0.02, 0.07, 0.0)
        assert an.format.CharHeight == 0.003 and an.format.Bold is True

        for mode in ("no_note", "no_annotation"):
            adapter.currentModel.note_mode = mode
            assert adapter._insert_sheet_note("X", 0, 0, 0.003, False) is None

        adapter.currentModel.note_mode = "no_format"
        plain = adapter._insert_sheet_note("X", 0.01, 0.02, 0.003, False)
        assert plain.format is None and plain.pos == (0.01, 0.02, 0.0)

    def test_place_note_moves_the_measured_box(self) -> None:
        adapter = _Harness()
        an = _Ann("ABCDE")  # 10 mm wide, 3 mm high
        box = adapter._place_note(an, "ABCDE", left=0.02, top=0.07)
        assert box == pytest.approx((0.02, 0.067, 0.03, 0.07))

        box = adapter._place_note(an, "ABCDE", centre_x=0.1, bottom=0.05)
        assert box == pytest.approx((0.095, 0.05, 0.105, 0.053))

        # Already in place: nothing moves.
        assert adapter._place_note(an, "ABCDE", centre_x=0.1) == pytest.approx(box)

        # A note whose position can't be read stays where it is.
        stuck = _Ann("AB")
        stuck.no_position = True
        assert adapter._place_note(stuck, "AB", left=0.5) == (0.0, -0.003, 0.004, 0.0)

    def test_note_box_falls_back_to_an_estimate(self) -> None:
        adapter = _Harness()
        box = adapter._note_box(_Ann("BROKEN"), "ABCD", 0.01, 0.05, 0.003)
        assert box == pytest.approx((0.01, 0.047, 0.01 + 0.62 * 0.003 * 4, 0.05))


# ---- the generator --------------------------------------------------------------


@pytest.fixture
def part(tmp_path, monkeypatch) -> str:
    """A part file next to an A3 sheet format the generator can find."""
    fmt_dir = tmp_path / "formats"
    fmt_dir.mkdir()
    (fmt_dir / "a3 - iso.slddrt").write_bytes(b"")
    monkeypatch.setattr(std.glob, "glob", lambda _pattern: [str(fmt_dir)])
    path = tmp_path / "Block.SLDPRT"
    path.write_bytes(b"")
    return str(path)


def _steps(result: AdapterResult[Any]) -> dict[str, str]:
    return {s["step"]: s["status"] for s in result.data["steps"]}


class TestCreateStandardDrawing:
    @pytest.mark.asyncio
    async def test_complete_drawing(self, part, tmp_path) -> None:
        adapter = _Harness()

        result = await adapter.create_standard_drawing(
            part,
            output_dir=str(tmp_path / "out"),
            title="Block",
            subtitle="Test block",
            drawing_number="BLK-01",
            author="saiq",
            company="Acme",
            drawn_date="25 Oct 2026",
            notes=["break edges", "  "],
        )

        assert result.is_success, result.error
        data = result.data
        assert data["scale"] == "1.5:1"
        assert data["isometric_scale"] is not None
        assert data["views"] == {
            "Drawing View1": {"role": "front", "caption": "FRONT ELEVATION"},
            "Drawing View2": {"role": "plan", "caption": "PLAN"},
            "Drawing View3": {"role": "side", "caption": "SIDE ELEVATION"},
            "Drawing View4": {"role": "iso", "caption": "ISOMETRIC"},
        }
        assert data["notes"] == [
            "NOTES",
            "1. MATERIAL: AISI 1020",
            "2. ALL DIMENSIONS IN MILLIMETRES",
            "3. THIRD ANGLE PROJECTION",
            "4. REMOVE ALL BURRS AND SHARP EDGES",
            "5. BREAK EDGES",
        ]
        assert data["overlap_check"] == {"overlaps_after": 0, "annotation_overlaps": []}
        assert data["dimensions"] == {"dimensions_added": 6}
        assert data["drawing_path"] == str(tmp_path / "out" / "BLK-01.SLDDRW")
        assert data["pdf_path"] == str(tmp_path / "out" / "BLK-01.pdf")
        assert data["drawing_number"] == "BLK-01"
        assert data["properties_written"]["drawing"]["COMPANYNAME"] == "Acme"
        assert set(_steps(result).values()) == {"success"}

        # The sheet-format prompt was off while the drawing was made, then restored.
        create = next(c for c in adapter.calls if c[0] == "create_drawing")
        assert create[1][std._SW_SHOW_SHEET_FORMAT_DIALOG] is False
        assert adapter.swApp.toggles[std._SW_SHOW_SHEET_FORMAT_DIALOG] is True
        # Dimensions only on the orthographic views; the final pass only checks.
        assert (
            "import_dimensions",
            ("Drawing View1", "Drawing View2", "Drawing View3"),
        ) in adapter.calls
        moves = [c[1] for c in adapter.calls if c[0] == "arrange_layout"]
        assert moves == [True, False]
        assert adapter.drawing.prefs[std._SW_DETAILING_DIM_TRAILING_ZERO] == 2
        assert adapter.drawing.sheet.scale == (1.5, 1.0)
        # Plan above the front, sharing its x; side beside it, sharing its y.
        views = adapter.drawing.views
        front, plan, side = (views[f"Drawing View{i}"].Position for i in (1, 2, 3))
        assert plan[0] == pytest.approx(front[0]) and plan[1] > front[1]
        assert side[1] == pytest.approx(front[1]) and side[0] > front[0]

    @pytest.mark.asyncio
    async def test_defaults_open_document_and_assembly(self, part, tmp_path) -> None:
        asm = tmp_path / "Frame Assy.SLDASM"
        asm.write_bytes(b"")
        adapter = _Harness()
        adapter.drawing.iso = False
        open_doc = SimpleNamespace(GetPathName=lambda: str(asm), GetTitle=lambda: "F")
        adapter.swApp.open_docs[str(asm)] = open_doc

        result = await adapter.create_standard_drawing(
            str(asm), revision="", include_isometric=True, write_model_properties=False
        )

        assert result.is_success, result.error
        assert result.data["drawing_number"] == "FRAME-ASSY-RevA-01"
        assert result.data["drawing_path"] == str(
            tmp_path / "FRAME-ASSY-RevA-01.SLDDRW"
        )
        assert result.data["isometric_scale"] is None
        assert "Drawing View4" not in result.data["views"]
        assert result.data["notes"][1] == "1. ALL DIMENSIONS IN MILLIMETRES"
        names = [c[0] for c in adapter.calls]
        assert "open_model" not in names and "get_material" not in names
        assert "model" not in result.data["properties_written"]

    @pytest.mark.asyncio
    async def test_input_errors(self, part, tmp_path) -> None:
        adapter = _Harness()
        missing = str(tmp_path / "nope.SLDPRT")
        drawing = tmp_path / "x.SLDDRW"
        drawing.write_bytes(b"")

        assert (await adapter.create_standard_drawing(missing)).error == (
            f"Model file not found: {missing}"
        )
        assert (await adapter.create_standard_drawing(str(drawing))).error == (
            "model_path must be a .SLDPRT or .SLDASM"
        )

    @pytest.mark.asyncio
    async def test_open_failure(self, part) -> None:
        adapter = _Harness()
        adapter.results["open_model"] = _err("locked")
        result = await adapter.create_standard_drawing(part)
        assert result.error.startswith("Could not open model: locked.")

    @pytest.mark.asyncio
    async def test_no_sheet_format(self, tmp_path, monkeypatch) -> None:
        monkeypatch.setattr(std.glob, "glob", lambda _pattern: [])
        path = tmp_path / "Block.SLDPRT"
        path.write_bytes(b"")
        adapter = _Harness()
        adapter.results["get_material"] = _err("no material")

        result = await adapter.create_standard_drawing(str(path))

        assert result.error == "Sheet format 'a3 - iso.slddrt' not found"

    @pytest.mark.asyncio
    async def test_drawing_creation_failure_restores_nothing_it_did_not_read(
        self, part
    ) -> None:
        adapter = _Harness()
        adapter.swApp.toggles = {}
        adapter.results["create_drawing"] = _err("no template")

        result = await adapter.create_standard_drawing(part)

        assert result.error == "Could not create drawing: no template"
        assert adapter.swApp.toggles == {std._SW_SHOW_SHEET_FORMAT_DIALOG: False}

    @pytest.mark.asyncio
    async def test_sheet_setup_failure(self, part) -> None:
        adapter = _Harness()
        adapter.drawing.setup_ok = False
        result = await adapter.create_standard_drawing(part)
        assert result.error.startswith(
            "Error in create_standard_drawing: SetupSheet5 failed"
        )

    @pytest.mark.asyncio
    async def test_no_views(self, part) -> None:
        adapter = _Harness()
        adapter.drawing.third_angle = "none"
        result = await adapter.create_standard_drawing(part)
        assert result.error.startswith(
            "Error in create_standard_drawing: No views could be created"
        )

    @pytest.mark.asyncio
    async def test_views_made_despite_a_false_return_and_no_front_name(
        self, part, monkeypatch
    ) -> None:
        adapter = _Harness()
        adapter.drawing.third_angle = "false_but_views"
        made = adapter.drawing.Create3rdAngleViews2

        def unnamed(path):
            ok = made(path)
            for view in adapter.drawing.views.values():
                view.orientation = ""
            adapter.drawing.views["Drawing View3"].outline = None
            return ok

        monkeypatch.setattr(adapter.drawing, "Create3rdAngleViews2", unnamed)

        result = await adapter.create_standard_drawing(part, include_isometric=False)

        assert result.is_success, result.error
        roles = {n: v["role"] for n, v in result.data["views"].items()}
        # A view whose outline can't be read gets no role (and no caption).
        assert roles == {
            "Drawing View1": "front",
            "Drawing View2": "plan",
            "Drawing View3": None,
        }

    @pytest.mark.asyncio
    async def test_notes_that_cannot_be_inserted_are_skipped(self, part) -> None:
        adapter = _Harness()
        adapter.drawing.note_mode = "no_note"
        result = await adapter.create_standard_drawing(part)
        assert result.is_success
        assert result.data["views"]["Drawing View1"]["caption"] == "FRONT ELEVATION"

    @pytest.mark.asyncio
    async def test_remaining_overlaps_and_failed_saves(self, part) -> None:
        adapter = _Harness()
        adapter.final_overlaps = 2
        adapter.results["save_file"] = _err("read-only")
        adapter.results["export_file"] = _err("no pdf")
        adapter.results["set_custom_property"] = _err("locked")
        adapter.results["set_units"] = _err("units")

        result = await adapter.create_standard_drawing(part, notes=["a"] * 9)

        assert result.error == (
            "Drawing saved but 2 overlap(s) remain - the no-overlap rule is not met"
        )
        assert result.data["drawing_path"] is None and result.data["pdf_path"] is None
        assert len(result.data["notes"]) == 8  # NOTES + at most 7 lines
        assert set(_steps(result).values()) == {"error"}
