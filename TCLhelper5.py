#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TCLhelper5.py — Python version of TCLhelper5 for Altair HyperMesh & HyperView (2022+)
Compatible with HyperWorks 2022, 2023, 2024+ (Desktop & Modern/Unified UI).

Features:
  - Full HyperMesh fastener grouping, solid centroid node creation, and fast O(N*M)
    CBUSH creation in pure Python.
  - Full HyperView result exporters (Displacements, Von Mises / Principal / Max Shear,
    Failure Index, Inter-Laminar Stress, Extreme Layer Stresses XX..YZ, ZX/YZ Shear,
    and 1D Fastener Forces).
  - Exports formatted Excel/LibreOffice Calc (.xlsx) workbooks with:
      * Self-contained visual clusters separated by narrow gap columns
      * Each criterion cluster shows [Allowable | Extreme Stress/Force | RF]
      * Merged 2-level colored cluster headers & medium box borders around each cluster
      * Right-aligned numbers & Reserve Factors (including '> 2')
      * Automatic red highlighting when Reserve Factor (RF) < 1.0 (or FI >= 1.0)
      * Envelope (Worst-Case / Max-Min) summary row at the bottom of every report
      * Auto-fitted column widths & frozen header rows
"""

import math
import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    _HAS_OPENPYXL = True
except Exception:
    _HAS_OPENPYXL = False

# ---------------------------------------------------------------------------
# 0. Altair HyperWorks (hm / hw / Tclinter) Bridge Setup (2022+ Compatible)
# ---------------------------------------------------------------------------
_HW_TCL = None
_HM_MODULE = None
_HM_ENTITIES = None
_HW_MODULE = None
_INSIDE_HYPERWORKS = False

try:
    import hw as _HW_MODULE

    _INSIDE_HYPERWORKS = True
except Exception:
    _HW_MODULE = None

try:
    from hw.common.tcl import Tclinter as _Tclinter

    _HW_TCL = _Tclinter.tcl
    _INSIDE_HYPERWORKS = True
except Exception:
    try:
        import Tclinter as _Tclinter

        _HW_TCL = _Tclinter.tcl
        _INSIDE_HYPERWORKS = True
    except Exception:
        _Tclinter = None
        _HW_TCL = None

try:
    import hm as _HM_MODULE
    import hm.entities as _HM_ENTITIES

    _INSIDE_HYPERWORKS = True
except Exception:
    _HM_MODULE = None
    _HM_ENTITIES = None


def eval_tcl(cmd: str) -> str:
    """
    Evaluate a Tcl command string inside HyperMesh / HyperView (2022+)
    and return the result as a string.
    """
    if _HW_TCL is not None:
        res = _HW_TCL.eval(cmd)
        return "" if res is None else str(res)
    if _HW_MODULE is not None and hasattr(_HW_MODULE, "evalTcl"):
        res = _HW_MODULE.evalTcl(cmd)
        return "" if res is None else str(res)
    if _HM_MODULE is not None:
        try:
            model = _HM_MODULE.Model()
            model.evaltclstring(cmd, 0)
            return ""
        except Exception:
            pass
    if tk._default_root is not None:
        try:
            res = tk._default_root.tk.eval(cmd)
            return "" if res is None else str(res)
        except Exception:
            pass
    raise RuntimeError(
        "HyperWorks (HyperMesh / HyperView) Tcl bridge is not available. "
        "Please run this script inside HyperMesh or HyperView 2022+."
    )


def _register_hv_tcl_helpers():
    """
    Register fast HWI Selection & Contour Query helpers in the embedded Tcl
    interpreter so element-iterator loops run at native speed without
    per-element Python<->Tcl marshaling overhead.
    """
    tcl_helpers = r"""
    proc ::pyhelper_ForceShowAll {} {
        foreach h {force_mod force_cli force_win force_page force_proj force_sess} {
            catch {$h ReleaseHandle}
        }
        hwi GetSessionHandle force_sess
        force_sess GetProjectHandle force_proj
        force_proj GetPageHandle force_page [force_proj GetActivePage]
        force_page GetWindowHandle force_win [force_page GetActiveWindow]
        force_win GetClientHandle force_cli
        force_cli GetModelHandle force_mod [force_cli GetActiveModel]

        force_mod UnMaskAll
        force_cli Draw

        foreach h {force_mod force_cli force_win force_page force_proj force_sess} {
            catch {$h ReleaseHandle}
        }
        update idletasks
        update
        after 50
    }

    proc ::pyhelper_ApplySingleElementSelection {mod_h elem_set_h sel_str} {
        set sel_str [string trim $sel_str]
        if {$sel_str eq ""} return

        # 1. Try direct HyperView syntax first
        if {![catch {eval "$elem_set_h Add \"$sel_str\""}]} {
            if {[$elem_set_h GetSize] > 0} return
        }

        # 2. Try direct component name: "component == $sel_str"
        if {![catch {eval "$elem_set_h Add \"component == $sel_str\""}]} {
            if {[$elem_set_h GetSize] > 0} return
        }

        # 2b. Try "comp 1-10" or "component 1-10"
        if {[regexp -nocase {^comp(?:onent)?\s+(.+)$} $sel_str -> comp_args]} {
            if {![catch {eval "$elem_set_h Add \"component $comp_args\""}]} {
                if {[$elem_set_h GetSize] > 0} return
            }
        }

        # 3. Check for wildcard pattern in component names (e.g. "Plaque*" or "*body*")
        if {[string match "*\\**" $sel_str] || [string match "*\\?*" $sel_str]} {
            catch {
                set temp_cid [$mod_h AddSelectionSet component]
                $mod_h GetSelectionSetHandle temp_csh $temp_cid
                $temp_csh Add all
                set all_cids [$temp_csh GetList]
                $temp_csh ReleaseHandle
                $mod_h RemoveSelectionSet $temp_cid

                foreach cid $all_cids {
                    catch {
                        $mod_h GetComponentHandle compH $cid
                        set cname [$compH GetLabel]
                        $compH ReleaseHandle
                        if {[string match -nocase $sel_str $cname]} {
                            catch {$elem_set_h Add "component $cid"}
                        }
                    }
                }
            }
            if {[$elem_set_h GetSize] > 0} return
        }

        # 4. Search Set by ID or Name (Model Sets and User Sets)
        set target $sel_str
        if {[regexp -nocase {(?:selectionset|set)\s*(?:==|\s)\s*(.+)} $sel_str -> match]} {
            set target [string trim $match " '\""]
        }

        set pools [list "SETS_ID_POOL" "User_Set" "" "set" "Model" "model"]
        catch {
            foreach t {set element node component model group} {
                catch {
                    foreach p [$mod_h GetPoolList $t] {
                        if {[lsearch -exact $pools $p] == -1} {
                            lappend pools $p
                        }
                    }
                }
            }
        }

        set found_elems [list]
        set matched_sid ""

        foreach p $pools {
            set sset_ids [list]
            if {$p eq ""} {
                catch {set sset_ids [$mod_h GetSelectionSetList]}
            } else {
                catch {set sset_ids [$mod_h GetSelectionSetList $p]}
            }

            foreach sid $sset_ids {
                catch {
                    set sh ""
                    if {$p eq ""} {
                        $mod_h GetSelectionSetHandle sh $sid
                    } else {
                        $mod_h GetSelectionSetHandle sh "$p $sid"
                    }

                    set s_id [$sh GetID]
                    set s_label [$sh GetLabel]

                    if {[string equal -nocase $target $s_id] || \
                        [string equal -nocase $target $s_label] || \
                        [string equal -nocase "Set $target" $s_label] || \
                        ([string match "*\\**" $target] && [string match -nocase $target $s_label])} {
                        set matched_sid $sid
                        set found_elems [$sh GetList]
                    }
                    $sh ReleaseHandle
                }
                if {[llength $found_elems] > 0} { break }
            }
            if {[llength $found_elems] > 0} { break }
        }

        if {$matched_sid ne "" || [llength $found_elems] > 0} {
            if {$matched_sid ne ""} {
                catch {$elem_set_h Add "selectionset == $matched_sid"}
            }
            if {[$elem_set_h GetSize] == 0 && [llength $found_elems] > 0} {
                set chunk_sz 300
                for {set i 0} {$i < [llength $found_elems]} {incr i $chunk_sz} {
                    set chunk [lrange $found_elems $i [expr {$i + $chunk_sz - 1}]]
                    if {[catch {$elem_set_h Add "id $chunk"}]} {
                        catch {$elem_set_h Add idlist [join $chunk ,]}
                    }
                }
            }
            if {[$elem_set_h GetSize] > 0} return
        }

        # 5. Numeric ID or range fallback (e.g. "1-100" -> "id 1-100")
        if {[regexp {^[0-9]+(-[0-9]+)?$} $sel_str]} {
            catch {$elem_set_h Add "id $sel_str"}
            if {[$elem_set_h GetSize] > 0} return
        }
    }

    proc ::pyhelper_ApplyElementSelection {mod_h elem_set_h sel_str} {
        set sel_str [string trim $sel_str]
        if {$sel_str eq ""} { set sel_str "all" }

        if {[string match "*,*" $sel_str]} {
            set items [split $sel_str ","]
            foreach itm $items {
                set itm [string trim $itm]
                if {$itm ne ""} {
                    ::pyhelper_ApplySingleElementSelection $mod_h $elem_set_h $itm
                }
            }
            if {[$elem_set_h GetSize] > 0} { return }
        } else {
            ::pyhelper_ApplySingleElementSelection $mod_h $elem_set_h $sel_str
            if {[$elem_set_h GetSize] > 0} { return }
        }

        if {[$elem_set_h GetSize] == 0} {
            catch {$elem_set_h Add all}
        }
    }

    # Query min and max contour value over current qry selection set
    proc ::pyhelper_QueryMinMax {qry_h} {
        set max_val ""
        set min_val ""
        $qry_h GetIteratorHandle iter
        for {iter First} {[iter Valid]} {iter Next} {
            set val [lindex [iter GetDataList] 1]
            if {[string is double -strict $val]} {
                if {$max_val eq "" || $val > $max_val} { set max_val $val }
                if {$min_val eq "" || $val < $min_val} { set min_val $val }
            }
        }
        iter ReleaseHandle
        return [list $max_val $min_val]
    }

    # Query per-element absolute extreme contour value for 1D forces
    proc ::pyhelper_QueryElementAbsExtreme {qry_h} {
        array unset elem_val
        $qry_h GetIteratorHandle iter
        for {iter First} {[iter Valid]} {iter Next} {
            set lst [iter GetDataList]
            set eid [lindex $lst 0]
            set val [lindex $lst 1]
            if {[string is double -strict $val]} {
                if {![info exists elem_val($eid)]} {
                    set elem_val($eid) $val
                } else {
                    if {abs($val) > abs($elem_val($eid))} {
                        set elem_val($eid) $val
                    }
                }
            }
        }
        iter ReleaseHandle
        set res [list]
        foreach eid [lsort -integer [array names elem_val]] {
            lappend res "$eid:$elem_val($eid)"
        }
        return [join $res " "]
    }

    # Resolve Set ID, Set Label or Component ID to human-readable Label/Name
    proc ::pyhelper_ResolveSetName {sel_str} {
        set sel_str [string trim $sel_str]
        if {$sel_str eq ""} { return "" }

        set target $sel_str
        if {[regexp -nocase {(?:selectionset|set)\s*(?:==|\s)\s*(.+)} $sel_str -> match]} {
            set target [string trim $match " '\""]
        }

        set resolved_name ""

        # HyperMesh native check
        catch {
            set hm_name [hm_getvalue sets id=$target dataname=name]
            if {$hm_name ne ""} { set resolved_name $hm_name }
        }
        if {$resolved_name ne ""} { return $resolved_name }

        catch {
            # Check existing global 'mod' or create session handles
            set mod_h ""
            set need_cleanup 0
            if {[info exists ::mod] && ![catch {$::mod GetID}]} {
                set mod_h $::mod
            } elseif {[info exists mod] && ![catch {$mod GetID}]} {
                set mod_h $mod
            } else {
                hwi GetSessionHandle r_sess
                r_sess GetProjectHandle r_proj
                r_proj GetPageHandle r_page [r_proj GetActivePage]
                r_page GetWindowHandle r_win [r_page GetActiveWindow]
                r_win GetClientHandle r_cli
                r_cli GetModelHandle r_mod [r_cli GetActiveModel]
                set mod_h r_mod
                set need_cleanup 1
            }

            # Method 1: Direct locator lookup in SETS_ID_POOL and User_Set
            foreach pool [list "SETS_ID_POOL" "User_Set" ""] {
                set loc $target
                if {$pool ne ""} { set loc "$pool $target" }
                catch {
                    $mod_h GetSelectionSetHandle r_sh $loc
                    set lbl ""
                    catch {set lbl [$r_sh GetLabel]}
                    if {$lbl eq ""} { catch {set lbl [$r_sh GetName]} }
                    $r_sh ReleaseHandle
                    if {$lbl ne ""} {
                        set resolved_name $lbl
                        break
                    }
                }
            }

            # Method 2: Iterate pools if direct locator didn't find it
            if {$resolved_name eq ""} {
                set pools [list "SETS_ID_POOL" "User_Set" "" "set" "Model" "model"]
                catch {
                    foreach t {selectionset set element node component model group} {
                        catch {
                            foreach p [$mod_h GetPoolList $t] {
                                if {[lsearch -exact $pools $p] == -1} {
                                    lappend pools $p
                                }
                            }
                        }
                    }
                }

                foreach p $pools {
                    set sset_ids [list]
                    if {$p eq ""} {
                        catch {set sset_ids [$mod_h GetSelectionSetList]}
                    } else {
                        catch {set sset_ids [$mod_h GetSelectionSetList $p]}
                    }

                    foreach sid $sset_ids {
                        catch {
                            set sh ""
                            if {$p eq ""} {
                                $mod_h GetSelectionSetHandle sh $sid
                            } else {
                                $mod_h GetSelectionSetHandle sh "$p $sid"
                            }

                            set s_id [$sh GetID]
                            set s_label ""
                            catch {set s_label [$sh GetLabel]}
                            if {$s_label eq ""} {
                                catch {set s_label [$sh GetName]}
                            }
                            $sh ReleaseHandle

                            if {[string equal -nocase $target $s_id] || \
                                [string equal -nocase $target $s_label] || \
                                [string equal -nocase "Set $target" $s_label]} {
                                if {$s_label ne ""} {
                                    set resolved_name $s_label
                                } else {
                                    set resolved_name $s_id
                                }
                                break
                            }
                        }
                    }
                    if {$resolved_name ne ""} { break }
                }
            }

            # Method 3: Components
            if {$resolved_name eq ""} {
                set comp_target $target
                if {[regexp -nocase {^comp(?:onent)?\s+(.+)$} $target -> c_arg]} {
                    set comp_target [string trim $c_arg]
                }
                catch {
                    set temp_cid [$mod_h AddSelectionSet component]
                    $mod_h GetSelectionSetHandle temp_csh $temp_cid
                    $temp_csh Add all
                    set all_cids [$temp_csh GetList]
                    $temp_csh ReleaseHandle
                    $mod_h RemoveSelectionSet $temp_cid

                    foreach cid $all_cids {
                        catch {
                            $mod_h GetComponentHandle compH $cid
                            set cname ""
                            catch {set cname [$compH GetLabel]}
                            if {$cname eq ""} { catch {set cname [$compH GetName]} }
                            $compH ReleaseHandle

                            if {[string equal -nocase $comp_target $cid] || \
                                [string equal -nocase $comp_target $cname]} {
                                set resolved_name $cname
                                break
                            }
                        }
                    }
                }
            }

            if {$need_cleanup} {
                foreach h {r_mod r_cli r_win r_page r_proj r_sess} {
                    catch {$h ReleaseHandle}
                }
            }
        }
        return $resolved_name
    }
    """
    try:
        eval_tcl(tcl_helpers)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 1. Pure Python Math, Formatting & Clustered Excel Report Engine
# ---------------------------------------------------------------------------
def _to_float(val):
    """Attempt to convert val to float; return None if empty or invalid."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val):
            return None
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    try:
        f = float(s.replace(",", "."))
        if math.isnan(f) or math.isinf(f):
            return None
        return f
    except ValueError:
        return None


def _clean_zero(fval: float, decimals: int = 2) -> float:
    """Avoid negative zero (-0.00) when rounding."""
    if fval is None:
        return None
    if abs(fval) < (10.0 ** (-decimals)) / 2.0:
        return 0.0
    return fval


def format_number(value, decimals: int = 2) -> str:
    """
    Format number to fixed decimals using comma ',' as decimal separator.
    Avoids negative zero (-0,00). Returns 'N/A' if invalid.
    """
    fval = _to_float(value)
    if fval is None:
        return "N/A"
    fval = _clean_zero(fval, decimals)
    formatted = f"{fval:.{decimals}f}"
    return formatted.replace(".", ",")


def safe_rf(allowable, stress) -> str:
    """
    Calculate Reserve Factor RF = |allowable / stress|.
    Returns 'N/A' if inputs are invalid, '> 2' if RF > 2.0 or stress ~ 0.
    """
    f_allow = _to_float(allowable)
    f_stress = _to_float(stress)
    if f_allow is None or f_stress is None:
        return "N/A"
    if f_stress == 0.0 or abs(f_stress) < 1e-12:
        return "> 2"
    rf_val = abs(f_allow / f_stress)
    if rf_val > 2.0:
        return "> 2"
    return format_number(rf_val, 2)


def get_extreme(maxval, minval) -> str:
    """Return formatted extreme value (whichever has the larger magnitude)."""
    f_max = _to_float(maxval)
    f_min = _to_float(minval)
    if f_max is None and f_min is None:
        return "N/A"
    if f_max is None:
        return format_number(f_min, 2)
    if f_min is None:
        return format_number(f_max, 2)
    if abs(f_min) > abs(f_max):
        return format_number(f_min, 2)
    return format_number(f_max, 2)


def get_extreme_float(maxval, minval):
    """Return numeric float extreme value (whichever has the larger magnitude)."""
    f_max = _to_float(maxval)
    f_min = _to_float(minval)
    if f_max is None and f_min is None:
        return None
    if f_max is None:
        return _clean_zero(f_min, 2)
    if f_min is None:
        return _clean_zero(f_max, 2)
    return _clean_zero(f_min if abs(f_min) > abs(f_max) else f_max, 2)


def _split_tcl_list(tcl_str: str) -> list:
    """Safely split a Tcl list string into a Python list of strings."""
    if not tcl_str or not tcl_str.strip():
        return []
    if _HW_TCL is not None:
        try:
            return list(_HW_TCL.splitlist(tcl_str))
        except Exception:
            pass
    if tk._default_root is not None:
        try:
            return list(tk._default_root.tk.splitlist(tcl_str))
        except Exception:
            pass
    tokens = []
    for match in re.finditer(r"\{([^}]*)\}|\"([^\"]*)\"|(\S+)", tcl_str):
        tokens.append(match.group(1) or match.group(2) or match.group(3))
    return tokens


def export_clustered_report(
    filepath: str,
    sheet_title: str,
    clusters: list,
    data_rows: list,
    rf_col_indices: set = None,
    fi_col_indices: set = None,
    summary_row: list = None,
) -> list:
    """
    Write a clean, clustered engineering report to:
      Formatted Excel/LibreOffice Calc (.xlsx) where each cluster is a visually
      separated block (with a narrow gap column between clusters, merged group
      header, medium border box, right-aligned numbers & '> 2', red highlight for
      RF < 1.0 or FI >= 1.0, and an Envelope summary row).

    Parameters
    ----------
    filepath : str
        Target path chosen by user (.xlsx).
    sheet_title : str
        Worksheet title for the .xlsx file.
    clusters : list of dicts
        Each dict defines a self-contained visual cluster:
          {
            "title": "VON MISES CRITERION",
            "color": "DDEBF7",      # Hex fill color for group header
            "sub_color": "EBF3FA",  # Hex fill color for column header
            "columns": ["Ftu", "Max_VonMises", "RF_VonMises"]
          }
    data_rows : list of lists
        Each row is a flat list of values matching the concatenated cluster columns.
    rf_col_indices : set of int (0-based into flat data_rows)
        Indices of Reserve Factor columns (highlighted red if RF < 1.0, green if '> 2').
    fi_col_indices : set of int (0-based into flat data_rows)
        Indices of Failure Index columns (highlighted red if FI >= 1.0).
    summary_row : list, optional
        Bottom summary / envelope row (same length as each data row).
    """
    if rf_col_indices is None:
        rf_col_indices = set()
    if fi_col_indices is None:
        fi_col_indices = set()

    base_path, ext = os.path.splitext(filepath)
    xlsx_path = filepath if ext.lower() == ".xlsx" else base_path + ".xlsx"
    saved_files = []

    if not _HAS_OPENPYXL:
        err_msg = (
            "The 'openpyxl' library is required to export .xlsx files.\n"
            "Please install openpyxl in your Python environment:\n"
            "    pip install openpyxl"
        )
        try:
            messagebox.showerror("Export Error", err_msg)
        except Exception:
            pass
        raise RuntimeError(err_msg)

    flat_headers = []
    for cl in clusters:
        flat_headers.extend(cl["columns"])

    # -----------------------------------------------------------------------
    # Write Formatted .xlsx Workbook (with narrow gap columns between clusters)
    # -----------------------------------------------------------------------
    if _HAS_OPENPYXL:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title[:31]
        ws.views.sheetView[0].showGridLines = True

        # Typography & Palette
        font_group = Font(name="Calibri", size=10, bold=True, color="1F2937")
        font_header = Font(name="Calibri", size=10, bold=True, color="111827")
        font_data = Font(name="Calibri", size=10, bold=False, color="111827")
        font_allow = Font(name="Calibri", size=10, italic=True, color="4B5563")
        font_bold = Font(name="Calibri", size=10, bold=True, color="111827")
        font_rf_fail = Font(name="Calibri", size=10, bold=True, color="9C0006")
        font_rf_ok = Font(name="Calibri", size=10, bold=True, color="276749")

        fill_rf_fail = PatternFill("solid", fgColor="FFC7CE")
        fill_rf_ok = PatternFill("solid", fgColor="E2EFDA")
        fill_zebra = PatternFill("solid", fgColor="F9FAFB")
        fill_white = PatternFill("solid", fgColor="FFFFFF")
        fill_allow = PatternFill("solid", fgColor="F3F4F6")
        fill_summary = PatternFill("solid", fgColor="E5E7EB")

        thin_side = Side(style="thin", color="D1D5DB")
        med_side = Side(style="medium", color="374151")
        double_side = Side(style="double", color="111827")

        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        # Map each flat data column index -> Excel 1-based column index
        # inserting 1 narrow blank gap column between adjacent clusters
        data_col_to_excel_col = {}
        gap_excel_cols = []
        cluster_excel_spans = []  # (start_excel_col, end_excel_col, cluster_dict)

        excel_col_cursor = 1
        flat_col_cursor = 0
        for c_idx, cl in enumerate(clusters):
            n_cols = len(cl["columns"])
            start_ec = excel_col_cursor
            end_ec = excel_col_cursor + n_cols - 1
            cluster_excel_spans.append((start_ec, end_ec, cl))

            for i in range(n_cols):
                data_col_to_excel_col[flat_col_cursor] = excel_col_cursor
                flat_col_cursor += 1
                excel_col_cursor += 1

            if c_idx < len(clusters) - 1:
                gap_excel_cols.append(excel_col_cursor)
                excel_col_cursor += 1

        # 1. Render Row 1 (Cluster Group Titles) & Row 2 (Column Headers)
        for start_ec, end_ec, cl in cluster_excel_spans:
            grp_fill = PatternFill("solid", fgColor=cl.get("color", "D9E1F2"))
            sub_fill = PatternFill("solid", fgColor=cl.get("sub_color", "EDF2F8"))

            if end_ec > start_ec:
                ws.merge_cells(
                    start_row=1,
                    start_column=start_ec,
                    end_row=1,
                    end_column=end_ec,
                )

            for c in range(start_ec, end_ec + 1):
                left_b = med_side if c == start_ec else thin_side
                right_b = med_side if c == end_ec else thin_side

                cell_g = ws.cell(row=1, column=c)
                if c == start_ec:
                    cell_g.value = cl["title"]
                cell_g.font = font_group
                cell_g.fill = grp_fill
                cell_g.alignment = align_center
                cell_g.border = Border(
                    left=left_b, right=right_b, top=med_side, bottom=thin_side
                )

                col_name = cl["columns"][c - start_ec]
                cell_h = ws.cell(row=2, column=c, value=col_name)
                cell_h.font = font_header
                cell_h.fill = sub_fill
                cell_h.alignment = align_center
                cell_h.border = Border(
                    left=left_b, right=right_b, top=thin_side, bottom=med_side
                )

        ws.row_dimensions[1].height = 24
        ws.row_dimensions[2].height = 22

        # Track cluster start/end in flat indices for border styling
        flat_cluster_starts = set()
        flat_cluster_ends = set()
        cur = 0
        for cl in clusters:
            flat_cluster_starts.add(cur)
            cur += len(cl["columns"])
            flat_cluster_ends.add(cur - 1)

        allowable_col_names = {
            "ft",
            "fs",
            "ftu",
            "fcy",
            "fcu",
            "fsu",
            "fzx",
            "fyz",
            "allowable_fi",
        }

        # 2. Render Data Rows + Optional Summary Row
        all_rows = list(data_rows)
        has_summary = summary_row is not None
        if has_summary:
            all_rows.append(summary_row)

        for r_idx, row_vals in enumerate(all_rows):
            excel_row = r_idx + 3
            is_sum_row = has_summary and (r_idx == len(all_rows) - 1)
            base_row_fill = (
                fill_summary
                if is_sum_row
                else (fill_zebra if r_idx % 2 == 1 else fill_white)
            )

            for flat_i in range(len(flat_headers)):
                ec = data_col_to_excel_col[flat_i]
                raw_val = row_vals[flat_i] if flat_i < len(row_vals) else ""
                cell = ws.cell(row=excel_row, column=ec)

                left_b = med_side if flat_i in flat_cluster_starts else thin_side
                right_b = med_side if flat_i in flat_cluster_ends else thin_side
                top_b = med_side if is_sum_row else thin_side
                bot_b = (
                    double_side
                    if is_sum_row
                    else (
                        med_side
                        if (not has_summary and r_idx == len(all_rows) - 1)
                        else thin_side
                    )
                )
                cell.border = Border(left=left_b, right=right_b, top=top_b, bottom=bot_b)

                col_name_lower = flat_headers[flat_i].lower()
                is_allow_col = col_name_lower in allowable_col_names

                cell.fill = (
                    fill_summary
                    if is_sum_row
                    else (fill_allow if is_allow_col else base_row_fill)
                )
                cell.font = (
                    font_bold
                    if is_sum_row
                    else (font_allow if is_allow_col else font_data)
                )

                # Value & Alignment ("align to droite" for all numbers & RF)
                if col_name_lower == "loadcase_name":
                    cell.value = str(raw_val)
                    cell.alignment = align_left
                elif col_name_lower in ("loadcase_id", "element_id"):
                    if isinstance(raw_val, int) or (
                        isinstance(raw_val, str) and raw_val.strip().isdigit()
                    ):
                        cell.value = int(raw_val)
                        cell.number_format = "0"
                    else:
                        cell.value = str(raw_val)
                    cell.alignment = align_right
                else:
                    cell.alignment = align_right
                    if isinstance(raw_val, (int, float)):
                        fval = _clean_zero(float(raw_val), 2)
                        cell.value = fval
                        cell.number_format = "0.00"
                    else:
                        s_val = str(raw_val).strip()
                        if s_val in ("> 2", "N/A", "-", ""):
                            cell.value = s_val if s_val else "N/A"
                        else:
                            f_parsed = _to_float(s_val)
                            if f_parsed is not None:
                                f_parsed = _clean_zero(f_parsed, 2)
                                cell.value = f_parsed
                                cell.number_format = "0.00"
                            else:
                                cell.value = s_val

                    # Conditional formatting for Reserve Factor (RF)
                    if flat_i in rf_col_indices:
                        if isinstance(cell.value, (int, float)) and cell.value < 1.0:
                            cell.fill = fill_rf_fail
                            cell.font = font_rf_fail
                        elif cell.value == "> 2":
                            cell.fill = fill_rf_ok
                            cell.font = font_rf_ok

                    # Conditional formatting for Failure Index (FI >= 1.0 fails)
                    if flat_i in fi_col_indices:
                        if isinstance(cell.value, (int, float)) and cell.value >= 1.0:
                            cell.fill = fill_rf_fail
                            cell.font = font_rf_fail

        # 3. Auto-fit Column Widths, Set Narrow Gap Columns & Freeze Header Panes
        ws.freeze_panes = "A3"
        for flat_i in range(len(flat_headers)):
            ec = data_col_to_excel_col[flat_i]
            col_letter = get_column_letter(ec)
            max_len = len(str(flat_headers[flat_i])) + 3
            for r in all_rows:
                if flat_i < len(r):
                    v = r[flat_i]
                    if isinstance(v, float):
                        s_len = len(f"{v:.2f}") + 2
                    else:
                        s_len = len(str(v)) + 2
                    if s_len > max_len:
                        max_len = s_len
            ws.column_dimensions[col_letter].width = max(11, min(max_len + 2, 42))

        for gap_ec in gap_excel_cols:
            ws.column_dimensions[get_column_letter(gap_ec)].width = 2.5

        wb.save(xlsx_path)
        saved_files.append(xlsx_path)

    return saved_files


# ---------------------------------------------------------------------------
# 2. Application UI & Logic Class
# ---------------------------------------------------------------------------
class TCLHelperApp:
    def __init__(self, master=None):
        _register_hv_tcl_helpers()

        # Bind to HyperWorks Tk root if running inside HyperWorks
        if _Tclinter is not None and tk._default_root is None:
            try:
                _Tclinter.Widget(".")
            except Exception:
                pass

        # Destroy existing .exportUI if re-running script
        try:
            eval_tcl("catch {destroy .exportUI}")
        except Exception:
            pass

        if master is not None:
            self.win = master
        elif tk._default_root is not None:
            self.win = tk.Toplevel(tk._default_root)
        else:
            self.win = tk.Tk()

        self.win.title("TCL Helper (Python 2022+)")
        self.win.geometry("450x900")
        try:
            self.win.attributes("-topmost", True)
        except Exception:
            pass

        self.help_win = None

        # Tk Variables (matching TCLhelper5 defaults)
        self.user_elem_sel = tk.StringVar(master=self.win, value="Mesh")

        self.ftu_stress = tk.StringVar(master=self.win, value="450.0")
        self.fcy_stress = tk.StringVar(master=self.win, value="450.0")
        self.fsu_stress = tk.StringVar(master=self.win, value="200.0")

        self.ftu = tk.StringVar(master=self.win, value="30.0")
        self.fcu = tk.StringVar(master=self.win, value="30.0")
        self.fsu = tk.StringVar(master=self.win, value="44.66")

        self.fzx = tk.StringVar(master=self.win, value="0.5")
        self.fyz = tk.StringVar(master=self.win, value="0.98")

        self.ft = tk.StringVar(master=self.win, value="1000.0")
        self.fs = tk.StringVar(master=self.win, value="1000.0")

        self._build_ui()

    def _build_ui(self):
        # Help Button
        btn_help = tk.Button(
            self.win,
            text="Instructions",
            command=self.show_help,
            bg="lightblue",
        )
        btn_help.pack(pady=(10, 10), padx=30, fill="x")

        # --- Section 1: HyperMesh ---
        lbl_hm = tk.Label(
            self.win, text="HyperMesh", font=("Helvetica", 10, "bold")
        )
        lbl_hm.pack(anchor="w", padx=1, pady=(0, 1))

        sep10 = tk.Frame(self.win, height=2, bd=1, relief="sunken")
        sep10.pack(fill="x", padx=10, pady=10)

        tk.Button(
            self.win,
            text="Group fasteners into an assembly",
            command=self.run_fastener_group,
        ).pack(pady=10, padx=30, fill="x")

        tk.Button(
            self.win,
            text="Create center node from solid fastener",
            command=self.run_node_create,
        ).pack(pady=10, padx=30, fill="x")

        tk.Button(
            self.win,
            text="Create CBUSH from two node lists",
            command=self.run_create_cbush,
        ).pack(pady=10, padx=30, fill="x")

        # --- Section 2: HyperView ---
        lbl_hv = tk.Label(
            self.win, text="HyperView", font=("Helvetica", 10, "bold")
        )
        lbl_hv.pack(anchor="w", padx=1, pady=(0, 1))

        sep11 = tk.Frame(self.win, height=2, bd=1, relief="sunken")
        sep11.pack(fill="x", padx=10, pady=10)

        frm_set = tk.Frame(self.win)
        frm_set.pack(pady=5)
        lbl_sel_text = (
            "Select Elements by:\n"
            "  - all            : All elements\n"
            "  - Mesh / 1001    : Set Name or ID\n"
            "  - Part1          : Component name\n"
            "  - comp 1-10      : Component IDs\n"
            "  - 1-100          : Element IDs"
        )
        tk.Label(frm_set, text=lbl_sel_text, justify="left").pack(
            side="left", padx=10
        )
        tk.Entry(frm_set, textvariable=self.user_elem_sel, width=25).pack(
            side="left", padx=5
        )

        # Part 1: Displacement Extreme Exporter
        tk.Button(
            self.win,
            text="Export Displacements (all)",
            command=self.export_displacement,
        ).pack(pady=5, padx=30, fill="x")

        # Part 2: Stress Exporter
        tk.Frame(self.win, height=2, bd=1, relief="sunken").pack(
            fill="x", padx=10, pady=10
        )
        frm_allow_stress = tk.Frame(self.win)
        frm_allow_stress.pack(pady=5)
        tk.Label(frm_allow_stress, text="Ftu").pack(side="left", padx=5)
        tk.Entry(frm_allow_stress, textvariable=self.ftu_stress, width=10).pack(
            side="left", padx=5
        )
        tk.Label(frm_allow_stress, text="Fcy").pack(side="left", padx=5)
        tk.Entry(frm_allow_stress, textvariable=self.fcy_stress, width=10).pack(
            side="left", padx=5
        )
        tk.Label(frm_allow_stress, text="Fsu").pack(side="left", padx=5)
        tk.Entry(frm_allow_stress, textvariable=self.fsu_stress, width=10).pack(
            side="left", padx=5
        )

        tk.Button(
            self.win,
            text="Stress VonMises, Max/Min Prin, Max Shear",
            command=self.export_stress_for_set,
        ).pack(pady=10, padx=30, fill="x")

        # Part 3: Failure Index Exporter
        tk.Frame(self.win, height=2, bd=1, relief="sunken").pack(
            fill="x", padx=10, pady=10
        )
        tk.Button(
            self.win,
            text="Failure Index",
            command=self.export_failure_index_for_set,
        ).pack(pady=10, padx=30, fill="x")

        # Part 4: Inter Laminar Stress Exporter
        tk.Frame(self.win, height=2, bd=1, relief="sunken").pack(
            fill="x", padx=10, pady=10
        )
        tk.Button(
            self.win,
            text="Stress Inter Laminar",
            command=self.export_inter_laminar_stress_for_set,
        ).pack(pady=10, padx=30, fill="x")

        # Part 5: Stress Extreme Layer Exporter
        tk.Frame(self.win, height=2, bd=1, relief="sunken").pack(
            fill="x", padx=10, pady=10
        )
        frm_allow = tk.Frame(self.win)
        frm_allow.pack(pady=5)
        tk.Label(frm_allow, text="Ftu").pack(side="left", padx=5)
        tk.Entry(frm_allow, textvariable=self.ftu, width=10).pack(
            side="left", padx=5
        )
        tk.Label(frm_allow, text="Fcu").pack(side="left", padx=5)
        tk.Entry(frm_allow, textvariable=self.fcu, width=10).pack(
            side="left", padx=5
        )
        tk.Label(frm_allow, text="Fsu").pack(side="left", padx=5)
        tk.Entry(frm_allow, textvariable=self.fsu, width=10).pack(
            side="left", padx=5
        )

        tk.Button(
            self.win,
            text="Stress XX, YY, ZZ, XY, ZX, YZ",
            command=self.export_stress_extreme_layer,
        ).pack(pady=10, padx=30, fill="x")

        # Part 6: Stress Extreme Layer Exporter (ZX & YZ Only)
        tk.Frame(self.win, height=2, bd=1, relief="sunken").pack(
            fill="x", padx=10, pady=10
        )
        frm_allow_zxyz = tk.Frame(self.win)
        frm_allow_zxyz.pack(pady=5)
        tk.Label(frm_allow_zxyz, text="Fzx").pack(side="left", padx=5)
        tk.Entry(frm_allow_zxyz, textvariable=self.fzx, width=10).pack(
            side="left", padx=5
        )
        tk.Label(frm_allow_zxyz, text="Fyz").pack(side="left", padx=5)
        tk.Entry(frm_allow_zxyz, textvariable=self.fyz, width=10).pack(
            side="left", padx=5
        )

        tk.Button(
            self.win,
            text="Stress ZX, YZ",
            command=self.export_stress_extreme_layer_zxyz,
        ).pack(pady=10, padx=30, fill="x")

        # Part 7: Fastener Force Exporter (1D Elements)
        tk.Frame(self.win, height=2, bd=1, relief="sunken").pack(
            fill="x", padx=10, pady=10
        )
        frm_fastener = tk.Frame(self.win)
        frm_fastener.pack(pady=5)
        tk.Label(frm_fastener, text="Ft").pack(side="left", padx=5)
        tk.Entry(frm_fastener, textvariable=self.ft, width=10).pack(
            side="left", padx=5
        )
        tk.Label(frm_fastener, text="Fs").pack(side="left", padx=5)
        tk.Entry(frm_fastener, textvariable=self.fs, width=10).pack(
            side="left", padx=5
        )

        tk.Button(
            self.win,
            text="1D Element Forces (X, Y, Z)",
            command=self.export_fastener_forces,
        ).pack(pady=10, padx=30, fill="x")

    # -----------------------------------------------------------------------
    # UI Dialog Helpers
    # -----------------------------------------------------------------------
    def _clean_set_name(self) -> str:
        try:
            raw = self.user_elem_sel.get().strip() if hasattr(self, "user_elem_sel") else ""
        except Exception:
            raw = ""
        if not raw:
            return "all"

        def _lookup_bdf_deck_set_name(target: str) -> str:
            """Fallback: search solver input decks (.bdf, .fem, .dat) for set/comp label."""
            try:
                tgt = target.strip()
                m = re.match(r'(?:selectionset|set)\s*(?:==|\s)\s*(.+)', tgt, re.IGNORECASE)
                if m:
                    tgt = m.group(1).strip(" '\"")
                search_dirs = [os.getcwd()]
                try:
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                    if script_dir not in search_dirs:
                        search_dirs.append(script_dir)
                except Exception:
                    pass

                for sdir in search_dirs:
                    if not os.path.isdir(sdir):
                        continue
                    for fname in os.listdir(sdir):
                        if fname.lower().endswith((".bdf", ".fem", ".dat")):
                            fpath = os.path.join(sdir, fname)
                            try:
                                with open(fpath, "r", encoding="utf-8", errors="ignore") as fp:
                                    for line in fp:
                                        if "$HMSET" in line:
                                            parts = line.split()
                                            if len(parts) >= 4 and parts[1].strip() == tgt:
                                                return parts[3].strip('"\'')
                                        if "$HMNAME COMP" in line or "$HMNAME SET" in line:
                                            parts = line.split()
                                            if len(parts) >= 4 and parts[2].strip() == tgt:
                                                return parts[3].strip('"\'')
                            except Exception:
                                pass
            except Exception:
                pass
            return ""

        def _resolve_single_token(tok: str) -> str:
            tok = tok.strip()
            if not tok:
                return ""
            # 1. Try Tcl resolution via HyperView / HyperMesh
            try:
                _register_hv_tcl_helpers()
                escaped = tok.replace("\\", "\\\\").replace('"', '\\"')
                tcl_res = eval_tcl(f'::pyhelper_ResolveSetName "{escaped}"')
                if tcl_res and str(tcl_res).strip():
                    return str(tcl_res).strip()
            except Exception:
                pass

            # 2. Try solver deck (.bdf / .fem) fallback in current working dir
            bdf_res = _lookup_bdf_deck_set_name(tok)
            if bdf_res:
                return bdf_res

            return tok

        if "," in raw:
            resolved_parts = [
                _resolve_single_token(part)
                for part in raw.split(",")
                if part.strip()
            ]
            target_name = "_".join(resolved_parts)
        else:
            target_name = _resolve_single_token(raw)

        # Sanitize characters forbidden in filenames: \ / : * ? " < > |
        cleaned = re.sub(r'[\\/*?:"<>|]', '_', target_name)
        # Collapse whitespace into single underscore
        cleaned = re.sub(r'\s+', '_', cleaned).strip('_')
        return cleaned or "all"

    def _ask_save_xlsx(self, title: str, default_name: str = "") -> str:
        if not default_name:
            default_name = f"Report_{self._clean_set_name()}.xlsx"
        elif not default_name.lower().endswith(".xlsx"):
            default_name += ".xlsx"

        return filedialog.asksaveasfilename(
            parent=self.win,
            title=title,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[
                ("Excel / LibreOffice Formatted (*.xlsx)", "*.xlsx"),
                ("All Files (*.*)", "*.*"),
            ],
        )

    _ask_save_csv = _ask_save_xlsx

    def show_help(self):
        if self.help_win is not None:
            try:
                self.help_win.destroy()
            except Exception:
                pass

        self.help_win = tk.Toplevel(self.win)
        self.help_win.title("Instructions")
        self.help_win.geometry("420x720")
        try:
            self.help_win.attributes("-topmost", True)
        except Exception:
            pass

        frm = tk.Frame(self.help_win)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        txt = tk.Text(
            frm,
            wrap="word",
            width=42,
            height=18,
            font=("Helvetica", 10),
            bg="#f0f0f0",
        )
        txt.pack(side="left", fill="both", expand=True)

        instructions = (
            "How to use this tool:\n\n"
            "HYPERMESH\n\n"
            "- Group fasteners into an assembly: groups all components whose name "
            "starts with the same reference (same NAS) into a single assembly.\n\n"
            "- Create center node from solid fastener: creates a center node for "
            "each solid fastener.\n\n"
            "- Create CBUSH from node lists: creates CBUSH elements between a list "
            "of node A and a list of node B, using the nearest B node for each A node.\n\n"
            "HYPERVIEW\n\n"
            "1. Load your model and all result files in HyperView.\n\n"
            "2. Selection Format:\n"
            "   - all            : All elements in model\n"
            "   - Mesh or 1001   : Set Name or Set ID\n"
            "   - Part1          : Component name\n"
            "   - comp 1-10      : Component ID range\n"
            "   - 1-100 or id 5  : Element IDs\n\n"
            "3. Enter your allowable values (Fsu, Fcu, Ftu, Ft, Fs, etc.) for Reserve Factor (RF) calculations.\n\n"
            "4. Click any export button. A formatted .xlsx workbook (clustered "
            "[Allowable | Value | RF] blocks, right-aligned numbers, red highlight for "
            "RF < 1.0, and Envelope summary row) is generated.\n"
        )

        txt.insert("end", instructions)
        txt.configure(state="disabled")

        tk.Button(
            self.help_win, text="Close", command=self.help_win.destroy
        ).pack(pady=10)

    # -----------------------------------------------------------------------
    # HYPERMESH PROCEDURES (2022+ Native Python API + Tcl Bridge)
    # -----------------------------------------------------------------------
    def run_fastener_group(self):
        base_to_ids = {}
        comp_found = False

        if _HM_MODULE is not None and _HM_ENTITIES is not None:
            try:
                model = _HM_MODULE.Model()
                comp_coll = _HM_MODULE.Collection(model, _HM_ENTITIES.Component)
                for comp in comp_coll:
                    comp_found = True
                    full_name = getattr(comp, "name", "") or ""
                    if not full_name.upper().startswith("NAS"):
                        continue
                    base_name = re.sub(r"\.\d+$", "", full_name)
                    base_to_ids.setdefault(base_name, []).append(int(comp.id))
            except Exception:
                comp_found = False
                base_to_ids.clear()

        if not comp_found:
            eval_tcl('*createmark comps 1 "all"')
            all_comp_ids = _split_tcl_list(eval_tcl("hm_getmark comps 1"))
            eval_tcl("hm_markclear comps 1")

            if not all_comp_ids:
                messagebox.showinfo(
                    "Info", "No components found in model.", parent=self.win
                )
                return

            for cid in all_comp_ids:
                full_name = eval_tcl(f"hm_entityinfo name components {cid}")
                if not full_name.upper().startswith("NAS"):
                    continue
                base_name = re.sub(r"\.\d+$", "", full_name)
                base_to_ids.setdefault(base_name, []).append(int(cid))

        count = 0
        for base_name, comp_ids in base_to_ids.items():
            id_str = " ".join(str(cid) for cid in comp_ids)
            eval_tcl(f'*createentity assems name="{base_name}"')
            assem_id = eval_tcl("hm_latestentityid assems").strip()
            eval_tcl(f"*createmark comps 1 {id_str}")
            eval_tcl(f"*assemblyaddmark {assem_id} comps 1")
            eval_tcl("hm_markclear comps 1")
            count += 1

        eval_tcl("hm_redraw")
        messagebox.showinfo(
            "Done",
            f"Fastener grouping complete.\n{count} assemblies created.",
            parent=self.win,
        )

    def run_node_create(self):
        eval_tcl("*clearmark solids 1")
        eval_tcl(
            '*createmarkpanel solids 1 "Select solids to create nodes at centroid:"'
        )
        solid_list = _split_tcl_list(eval_tcl("hm_getmark solids 1"))
        eval_tcl("*clearmark solids 1")

        if not solid_list:
            print("No solids selected.")
            return

        node_count = 0
        for solid_id in solid_list:
            eval_tcl(f"*createmark solids 1 {solid_id}")
            centroids = _split_tcl_list(eval_tcl("hm_getcentroid solids 1"))
            eval_tcl("*clearmark solids 1")
            if len(centroids) >= 3:
                x, y, z = centroids[0], centroids[1], centroids[2]
                eval_tcl(f"*createnode {x} {y} {z} 0 0 0")
                node_count += 1

        print(f"Successfully created {node_count} free node(s) at solid centroids.")
        messagebox.showinfo(
            "Done",
            f"Created {node_count} node(s) at solid centroids.",
            parent=self.win,
        )

    def _get_node_coordinates(self, node_ids: list) -> dict:
        coords = {}
        if not node_ids:
            return coords

        if _HM_MODULE is not None and _HM_ENTITIES is not None:
            try:
                model = _HM_MODULE.Model()
                for nid in node_ids:
                    node_obj = model.get(
                        _HM_MODULE.Entity(_HM_ENTITIES.Node, int(nid))
                    )
                    if node_obj is not None:
                        gcoords = getattr(node_obj, "globalcoordinates", None)
                        if gcoords is not None and len(gcoords) >= 3:
                            coords[str(nid)] = (
                                float(gcoords[0]),
                                float(gcoords[1]),
                                float(gcoords[2]),
                            )
                        elif hasattr(node_obj, "x"):
                            coords[str(nid)] = (
                                float(node_obj.x),
                                float(node_obj.y),
                                float(node_obj.z),
                            )
                if len(coords) == len(node_ids):
                    return coords
            except Exception:
                coords.clear()

        try:
            id_list_str = " ".join(str(nid) for nid in node_ids)
            raw = eval_tcl(
                f"""
                set _res [list]
                foreach _nid [list {id_list_str}] {{
                    set _x [hm_getvalue nodes id=$_nid dataname=x]
                    set _y [hm_getvalue nodes id=$_nid dataname=y]
                    set _z [hm_getvalue nodes id=$_nid dataname=z]
                    lappend _res "$_nid:$_x:$_y:$_z"
                }}
                join $_res " "
                """
            )
            for item in raw.split():
                parts = item.split(":")
                if len(parts) == 4:
                    coords[parts[0]] = (
                        float(parts[1]),
                        float(parts[2]),
                        float(parts[3]),
                    )
        except Exception:
            coords.clear()

        return coords

    def run_create_cbush(self):
        eval_tcl("*clearmark nodes 1")
        eval_tcl('*createmarkpanel nodes 1 "Select all Nodes A:"')
        nodes_a = _split_tcl_list(eval_tcl("hm_getmark nodes 1"))
        eval_tcl("*clearmark nodes 1")

        if not nodes_a:
            messagebox.showerror(
                "Error", "No nodes selected for List A.", parent=self.win
            )
            return

        eval_tcl('*createmarkpanel nodes 1 "Select all Nodes B:"')
        nodes_b = _split_tcl_list(eval_tcl("hm_getmark nodes 1"))
        eval_tcl("*clearmark nodes 1")

        if not nodes_b:
            messagebox.showerror(
                "Error", "No nodes selected for List B.", parent=self.win
            )
            return

        eval_tcl("*clearmark properties 1")
        eval_tcl(
            '*createmarkpanel properties 1 "Select the PBUSH property for the new elements:"'
        )
        prop_ids = _split_tcl_list(eval_tcl("hm_getmark properties 1"))
        eval_tcl("*clearmark properties 1")

        if not prop_ids:
            messagebox.showerror(
                "Error",
                "No property selected. Operation cancelled.",
                parent=self.win,
            )
            return

        selected_prop_id = prop_ids[0]
        eval_tcl("*elementtype spring CBUSH")

        cbush_count = 0
        new_elements_list = []

        eval_tcl("catch { *startnotehistorystate {Created CBUSH elements} }")

        all_unique_nodes = list(dict.fromkeys(nodes_a + nodes_b))
        coords = self._get_node_coordinates(all_unique_nodes)
        use_fast_coords = len(coords) == len(all_unique_nodes)

        for node_a in nodes_a:
            shortest_dist_sq = float("inf")
            nearest_node_b = ""

            if use_fast_coords:
                ax, ay, az = coords[str(node_a)]
                for node_b in nodes_b:
                    if str(node_a) == str(node_b):
                        continue
                    bx, by, bz = coords[str(node_b)]
                    dist_sq = (ax - bx) ** 2 + (ay - by) ** 2 + (az - bz) ** 2
                    if dist_sq < shortest_dist_sq:
                        shortest_dist_sq = dist_sq
                        nearest_node_b = str(node_b)
            else:
                for node_b in nodes_b:
                    if str(node_a) == str(node_b):
                        continue
                    dist_info = _split_tcl_list(
                        eval_tcl(f"hm_getdistance nodes {node_a} {node_b} 0")
                    )
                    if dist_info:
                        dist = float(dist_info[0])
                        if dist < shortest_dist_sq:
                            shortest_dist_sq = dist
                            nearest_node_b = str(node_b)

            if nearest_node_b:
                err_flag = eval_tcl(
                    f'catch {{*springos {node_a} {nearest_node_b} "" 0 0 0 0 1 1 0}} _cbush_err'
                ).strip()
                if err_flag == "0":
                    cbush_count += 1
                    new_elem_id = eval_tcl("hm_latestentityid elements").strip()
                    if new_elem_id:
                        new_elements_list.append(new_elem_id)
                else:
                    err_msg = eval_tcl("set _cbush_err")
                    print(
                        f"Failed to create CBUSH between {node_a} and {nearest_node_b}. Error: {err_msg}"
                    )

        if new_elements_list:
            elem_str = " ".join(new_elements_list)
            eval_tcl("*clearmark elements 1")
            eval_tcl(f"*createmark elements 1 {elem_str}")
            eval_tcl(f"*setvalue elements mark=1 propertyid={selected_prop_id}")
            eval_tcl("*clearmark elements 1")

        eval_tcl("catch { *endnotehistorystate {Created CBUSH elements} }")
        eval_tcl("hm_redraw")

        messagebox.showinfo(
            "Success",
            f"Created {cbush_count} CBUSH elements and assigned PBUSH property ID: {selected_prop_id}!",
            parent=self.win,
        )

    # -----------------------------------------------------------------------
    # HYPERVIEW PROCEDURES (2022+ Compatible, Clustered .xlsx Export)
    # -----------------------------------------------------------------------
    def _hv_init_handles(self, include_query: bool = False) -> int:
        _register_hv_tcl_helpers()
        eval_tcl(
            "foreach h {sess proj page win cli mod res cont leg qry iter elem_set} { catch {$h ReleaseHandle} }"
        )
        eval_tcl(
            """
            hwi GetSessionHandle sess
            sess GetProjectHandle proj
            proj GetPageHandle page [proj GetActivePage]
            page GetWindowHandle win [page GetActiveWindow]
            win GetClientHandle cli
            cli GetModelHandle mod [cli GetActiveModel]
            ::pyhelper_ForceShowAll
            mod GetResultCtrlHandle res
            res GetContourCtrlHandle cont
            """
        )
        if not include_query:
            return -1

        sel_str = self.user_elem_sel.get()
        escaped_sel = sel_str.replace("\\", "\\\\").replace('"', '\\"')
        num_elems_str = eval_tcl(
            f"""
            set ::pyhelper_sel_id [mod AddSelectionSet element]
            mod GetSelectionSetHandle elem_set $::pyhelper_sel_id
            ::pyhelper_ApplyElementSelection mod elem_set "{escaped_sel}"
            set _num [elem_set GetSize]
            mod GetQueryCtrlHandle qry
            qry SetSelectionSet $::pyhelper_sel_id
            qry SetQuery "element.id contour.value"
            return $_num
            """
        )
        try:
            return int(num_elems_str.strip())
        except ValueError:
            return 0

    def _hv_cleanup_handles(self, has_selection_set: bool = False):
        if has_selection_set:
            eval_tcl("catch {mod RemoveSelectionSet $::pyhelper_sel_id}")
        eval_tcl(
            "foreach h {cont res mod cli win page proj sess qry elem_set iter leg} { catch {$h ReleaseHandle} }"
        )

    # --- HV Part 1: Export Displacements (all) ---
    def export_displacement(self):
        try:
            self._hv_init_handles(include_query=False)
            filepath = self._ask_save_xlsx(
                "Save Extreme Displacements",
                default_name=f"Displacement_{self._clean_set_name()}.xlsx",
            )
            if not filepath:
                return

            subcase_list = _split_tcl_list(eval_tcl('res GetSubcaseList "Base"'))
            if not subcase_list:
                messagebox.showwarning(
                    "Warning", "No subcases found in active model.", parent=self.win
                )
                return

            first_sub = subcase_list[0]
            dtype = "Displacements"
            all_types = _split_tcl_list(
                eval_tcl(
                    f"if {{![catch {{res GetDataTypeList {first_sub}}} _t]}} {{return $_t}} else {{return {{}}}}"
                )
            )
            for t in all_types:
                if "displacement" in t.lower():
                    dtype = t
                    break

            eval_tcl(
                f"""
                cont SetEnableState false
                catch {{cont SetDataType "Displacement"}}
                catch {{cont SetDataType "{dtype}"}}
                cont SetEnableState true
                cli Draw
                """
            )

            comp_list = _split_tcl_list(
                eval_tcl(f'res GetDataComponentList {first_sub} "{dtype}"')
            )
            mag_str, x_str, y_str, z_str = "Mag", "X", "Y", "Z"
            for comp in comp_list:
                c_low = comp.lower()
                if c_low == "mag":
                    mag_str = comp
                elif c_low == "x":
                    x_str = comp
                elif c_low == "y":
                    y_str = comp
                elif c_low == "z":
                    z_str = comp

            rows = []
            g_x = g_y = g_z = g_mag = None

            for sub in subcase_list:
                eval_tcl(f"res SetCurrentSubcase {sub}")
                sub_name = eval_tcl(f"res GetSubcaseLabel {sub}").replace(",", "-")

                def _get_legend_min_max(comp_name: str):
                    raw = eval_tcl(
                        f"""
                        cont SetEnableState false
                        catch {{cont SetDataType "{dtype}"}}
                        cont SetDataComponent "{comp_name}"
                        cont SetEnableState true
                        cli Draw
                        cont GetLegendHandle leg
                        set _mx [leg GetMax]
                        set _mn [leg GetMin]
                        leg ReleaseHandle
                        list $_mx $_mn
                        """
                    )
                    parts = _split_tcl_list(raw)
                    return (
                        parts[0] if len(parts) > 0 else "",
                        parts[1] if len(parts) > 1 else "",
                    )

                mag_max_raw, _ = _get_legend_min_max(mag_str)
                mag_max = _to_float(mag_max_raw)

                x_max, x_min = _get_legend_min_max(x_str)
                x_ext = get_extreme_float(x_max, x_min)

                y_max, y_min = _get_legend_min_max(y_str)
                y_ext = get_extreme_float(y_max, y_min)

                z_max, z_min = _get_legend_min_max(z_str)
                z_ext = get_extreme_float(z_max, z_min)

                if x_ext is not None and (g_x is None or abs(x_ext) > abs(g_x)):
                    g_x = x_ext
                if y_ext is not None and (g_y is None or abs(y_ext) > abs(g_y)):
                    g_y = y_ext
                if z_ext is not None and (g_z is None or abs(z_ext) > abs(g_z)):
                    g_z = z_ext
                if mag_max is not None and (g_mag is None or mag_max > g_mag):
                    g_mag = mag_max

                rows.append([sub, sub_name, x_ext, y_ext, z_ext, mag_max])

            summary_row = [
                "MAX",
                "ENVELOPE (WORST CASE)",
                g_x,
                g_y,
                g_z,
                g_mag,
            ]

            clusters = [
                {
                    "title": "IDENTIFICATION",
                    "color": "D9E1F2",
                    "sub_color": "EDF2F8",
                    "columns": ["Loadcase_ID", "Loadcase_Name"],
                },
                {
                    "title": "DIRECTIONAL EXTREMES (X, Y, Z)",
                    "color": "DDEBF7",
                    "sub_color": "EBF3FA",
                    "columns": ["X_Extreme", "Y_Extreme", "Z_Extreme"],
                },
                {
                    "title": "RESULTANT MAGNITUDE",
                    "color": "FCE4D6",
                    "sub_color": "FDF0E6",
                    "columns": ["Mag_Max"],
                },
            ]
            saved = export_clustered_report(
                filepath,
                "Displacements",
                clusters,
                rows,
                summary_row=summary_row,
            )
            messagebox.showinfo(
                "Export Complete",
                "Export complete:\n" + "\n".join(saved),
                parent=self.win,
            )
        finally:
            self._hv_cleanup_handles(has_selection_set=False)

    # --- HV Part 2: Export Stress (VonMises, Max/Min Principal, Max Shear) ---
    # Each criterion is its own self-contained cluster: [Allowable | Stress | RF]
    def export_stress_for_set(self):
        ftu_str = self.ftu_stress.get()
        fcy_str = self.fcy_stress.get()
        fsu_str = self.fsu_stress.get()
        ftu_val = _to_float(ftu_str)
        fcy_val = _to_float(fcy_str)
        fsu_val = _to_float(fsu_str)
        # Fallback if Fcy is left empty: use Ftu
        comp_allow = fcy_val if fcy_val is not None else ftu_val

        try:
            num_elems = self._hv_init_handles(include_query=True)
            if num_elems == 0:
                messagebox.showwarning(
                    "Warning",
                    f"Warning: Your selection '{self.user_elem_sel.get()}' resulted in 0 elements.",
                    parent=self.win,
                )
                return

            filepath = self._ask_save_xlsx(
                "Save Extreme Stresses",
                default_name=f"Stress_VonMises_{self._clean_set_name()}.xlsx",
            )
            if not filepath:
                return

            subcase_list = _split_tcl_list(eval_tcl('res GetSubcaseList "Base"'))
            if not subcase_list:
                return
            first_sub = subcase_list[0]

            comp_list = _split_tcl_list(
                eval_tcl(f'res GetDataComponentList {first_sub} "Stresses"')
            )
            if not comp_list:
                comp_list = _split_tcl_list(
                    eval_tcl(f'res GetDataComponentList {first_sub} "Stress"')
                )

            c_von = "vonMises"
            c_max = "maxPrincipal"
            c_min = "minPrincipal"
            c_shear = "MaxShear"

            for c in comp_list:
                cl = c.lower()
                if re.search(r"von.*mises", cl) and "signed" not in cl:
                    c_von = c
                if re.search(r"max.*principal", cl) or (
                    re.search(r"p1.*major", cl) and not re.search(r"in.*plane", cl)
                ):
                    c_max = c
                if re.search(r"min.*principal", cl) or (
                    re.search(r"p3.*minor", cl) and not re.search(r"in.*plane", cl)
                ):
                    c_min = c
                if "maxshear" in cl:
                    c_shear = c

            rows = []
            g_max_von = g_max_p1 = g_min_p3 = g_max_shear = None

            for sub in subcase_list:
                eval_tcl(f"res SetCurrentSubcase {sub}")
                sub_name = eval_tcl(f"res GetSubcaseLabel {sub}").replace(",", "-")

                max_von = max_p1 = min_p3 = max_shear = None

                if num_elems > 0:
                    res_von = _split_tcl_list(
                        eval_tcl(
                            f"""
                            cont SetEnableState false
                            catch {{cont SetDataType "Stress"}}
                            catch {{cont SetDataComponent "{c_von}"}}
                            catch {{cont SetShellLayer "Max"}}
                            cont SetEnableState true
                            cli Draw
                            ::pyhelper_QueryMinMax qry
                            """
                        )
                    )
                    max_von = _to_float(res_von[0]) if res_von else None

                    res_p1 = _split_tcl_list(
                        eval_tcl(
                            f"""
                            cont SetEnableState false
                            catch {{cont SetDataComponent "{c_max}"}}
                            catch {{cont SetShellLayer "Max"}}
                            cont SetEnableState true
                            cli Draw
                            ::pyhelper_QueryMinMax qry
                            """
                        )
                    )
                    max_p1 = _to_float(res_p1[0]) if res_p1 else None

                    res_p3 = _split_tcl_list(
                        eval_tcl(
                            f"""
                            cont SetEnableState false
                            catch {{cont SetDataComponent "{c_min}"}}
                            catch {{cont SetShellLayer "Min"}}
                            cont SetEnableState true
                            cli Draw
                            ::pyhelper_QueryMinMax qry
                            """
                        )
                    )
                    min_p3 = _to_float(res_p3[1]) if len(res_p3) > 1 else None

                    res_sh = _split_tcl_list(
                        eval_tcl(
                            f"""
                            cont SetEnableState false
                            catch {{cont SetDataComponent "{c_shear}"}}
                            catch {{cont SetShellLayer "Max"}}
                            cont SetEnableState true
                            cli Draw
                            ::pyhelper_QueryMinMax qry
                            """
                        )
                    )
                    max_shear = _to_float(res_sh[0]) if res_sh else None

                if max_von is not None and (g_max_von is None or max_von > g_max_von):
                    g_max_von = max_von
                if max_p1 is not None and (g_max_p1 is None or max_p1 > g_max_p1):
                    g_max_p1 = max_p1
                if min_p3 is not None and (g_min_p3 is None or min_p3 < g_min_p3):
                    g_min_p3 = min_p3
                if max_shear is not None and (
                    g_max_shear is None or max_shear > g_max_shear
                ):
                    g_max_shear = max_shear

                rows.append(
                    [
                        sub,
                        sub_name,
                        # Cluster 1: Von Mises [Ftu, Max_VonMises, RF_VonMises]
                        ftu_val if ftu_val is not None else ftu_str,
                        max_von,
                        safe_rf(ftu_val, max_von),
                        # Cluster 2: Max Principal [Ftu, Max_MaxPrincipal, RF_MaxPrincipal]
                        ftu_val if ftu_val is not None else ftu_str,
                        max_p1,
                        safe_rf(ftu_val, max_p1),
                        # Cluster 3: Min Principal [Fcy, Min_MinPrincipal, RF_MinPrincipal]
                        comp_allow if comp_allow is not None else fcy_str,
                        min_p3,
                        safe_rf(comp_allow, min_p3),
                        # Cluster 4: Max Shear [Fsu, MaxShear, RF_MaxShear]
                        fsu_val if fsu_val is not None else fsu_str,
                        max_shear,
                        safe_rf(fsu_val, max_shear),
                    ]
                )

            summary_row = [
                "MAX/MIN",
                "ENVELOPE (WORST CASE)",
                ftu_val if ftu_val is not None else ftu_str,
                g_max_von,
                safe_rf(ftu_val, g_max_von),
                ftu_val if ftu_val is not None else ftu_str,
                g_max_p1,
                safe_rf(ftu_val, g_max_p1),
                comp_allow if comp_allow is not None else fcy_str,
                g_min_p3,
                safe_rf(comp_allow, g_min_p3),
                fsu_val if fsu_val is not None else fsu_str,
                g_max_shear,
                safe_rf(fsu_val, g_max_shear),
            ]

            clusters = [
                {
                    "title": "IDENTIFICATION",
                    "color": "D9E1F2",
                    "sub_color": "EDF2F8",
                    "columns": ["Loadcase_ID", "Loadcase_Name"],
                },
                {
                    "title": "VON MISES CRITERION",
                    "color": "DDEBF7",
                    "sub_color": "EBF3FA",
                    "columns": ["Ftu", "Max_VonMises", "RF_VonMises"],
                },
                {
                    "title": "MAX PRINCIPAL (TENSION)",
                    "color": "E2EFDA",
                    "sub_color": "EFF7EB",
                    "columns": ["Ftu", "Max_MaxPrincipal", "RF_MaxPrincipal"],
                },
                {
                    "title": "MIN PRINCIPAL (COMPRESSION)",
                    "color": "FFF2CC",
                    "sub_color": "FFF8E5",
                    "columns": ["Fcy", "Min_MinPrincipal", "RF_MinPrincipal"],
                },
                {
                    "title": "MAX SHEAR CRITERION",
                    "color": "FCE4D6",
                    "sub_color": "FDF0E6",
                    "columns": ["Fsu", "MaxShear", "RF_MaxShear"],
                },
            ]

            saved = export_clustered_report(
                filepath,
                "VonMises_Principal_Shear",
                clusters,
                rows,
                rf_col_indices={4, 7, 10, 13},
                summary_row=summary_row,
            )
            messagebox.showinfo(
                "Export Complete",
                "Stress Export complete!\nSaved to:\n" + "\n".join(saved),
                parent=self.win,
            )
        finally:
            self._hv_cleanup_handles(has_selection_set=True)

    # --- HV Part 3: Export Failure Index ---
    def export_failure_index_for_set(self):
        try:
            num_elems = self._hv_init_handles(include_query=True)
            if num_elems == 0:
                messagebox.showwarning(
                    "Warning",
                    "Warning: Your selection resulted in 0 elements.",
                    parent=self.win,
                )
                return

            filepath = self._ask_save_xlsx(
                "Save Extreme Failure Indices",
                default_name=f"Failure_Index_{self._clean_set_name()}.xlsx",
            )
            if not filepath:
                return

            subcase_list = _split_tcl_list(eval_tcl('res GetSubcaseList "Base"'))
            if not subcase_list:
                return
            first_sub = subcase_list[0]

            actual_type = "Failure Index (s)"
            actual_comp = "for direct Stress"

            all_types = _split_tcl_list(
                eval_tcl(
                    f"if {{![catch {{res GetDataTypeList {first_sub}}} _t]}} {{return $_t}} else {{return {{}}}}"
                )
            )
            for t in all_types:
                if re.search(r"failure.*index", t, re.IGNORECASE):
                    actual_type = t
                    break

            all_comps = _split_tcl_list(
                eval_tcl(
                    f'if {{![catch {{res GetDataComponentList {first_sub} "{actual_type}"}} _c]}} {{return $_c}} else {{return {{}}}}'
                )
            )
            for c in all_comps:
                if re.search(r"direct.*stress", c, re.IGNORECASE):
                    actual_comp = c
                    break

            def _calc_fi_rf(fi_val):
                if fi_val is None:
                    return "N/A"
                if fi_val > 0.0:
                    rf_val = 1.0 / math.sqrt(fi_val)
                    return "> 2" if rf_val > 2.0 else format_number(rf_val, 2)
                return "> 2"

            rows = []
            g_max_fi = None

            for sub in subcase_list:
                eval_tcl(f"res SetCurrentSubcase {sub}")
                sub_name = eval_tcl(f"res GetSubcaseLabel {sub}").replace(",", "-")

                max_fi = None
                if num_elems > 0:
                    res_fi = _split_tcl_list(
                        eval_tcl(
                            f"""
                            cont SetEnableState false
                            cont SetDataType "{actual_type}"
                            cont SetDataComponent "{actual_comp}"
                            cont SetShellLayer "Max"
                            cont SetEnableState true
                            cli Draw
                            ::pyhelper_QueryMinMax qry
                            """
                        )
                    )
                    max_fi = _to_float(res_fi[0]) if res_fi else None

                if max_fi is not None and (g_max_fi is None or max_fi > g_max_fi):
                    g_max_fi = max_fi

                rows.append([sub, sub_name, 1.0, max_fi, _calc_fi_rf(max_fi)])

            summary_row = [
                "MAX",
                "ENVELOPE (WORST CASE)",
                1.0,
                g_max_fi,
                _calc_fi_rf(g_max_fi),
            ]

            clusters = [
                {
                    "title": "IDENTIFICATION",
                    "color": "D9E1F2",
                    "sub_color": "EDF2F8",
                    "columns": ["Loadcase_ID", "Loadcase_Name"],
                },
                {
                    "title": "COMPOSITE FAILURE INDEX CRITERION",
                    "color": "FCE4D6",
                    "sub_color": "FDF0E6",
                    "columns": ["Allowable_FI", "Max_Failure_Index", "RF"],
                },
            ]
            saved = export_clustered_report(
                filepath,
                "Failure_Index",
                clusters,
                rows,
                rf_col_indices={4},
                fi_col_indices={3},
                summary_row=summary_row,
            )
            messagebox.showinfo(
                "Export Complete",
                "Failure Index Export complete!\nSaved to:\n"
                + "\n".join(saved)
                + f"\n\nDebug Info:\nType Used: '{actual_type}'\nComp Used: '{actual_comp}'",
                parent=self.win,
            )
        finally:
            self._hv_cleanup_handles(has_selection_set=True)

    # --- HV Part 4: Export Inter Laminar Stress ---
    def export_inter_laminar_stress_for_set(self):
        try:
            num_elems = self._hv_init_handles(include_query=True)
            if num_elems == 0:
                messagebox.showwarning(
                    "Warning",
                    "Warning: Your selection resulted in 0 elements.",
                    parent=self.win,
                )
                return

            filepath = self._ask_save_xlsx(
                "Save Extreme Inter Laminar Stresses",
                default_name=f"InterLaminar_Stress_{self._clean_set_name()}.xlsx",
            )
            if not filepath:
                return

            subcase_list = _split_tcl_list(eval_tcl('res GetSubcaseList "Base"'))
            if not subcase_list:
                return
            first_sub = subcase_list[0]

            actual_type = "Inter Laminar Stress (s)"
            comp_yz = "Shear YZ"
            comp_xz = "Shear XZ"

            all_types = _split_tcl_list(
                eval_tcl(
                    f"if {{![catch {{res GetDataTypeList {first_sub}}} _t]}} {{return $_t}} else {{return {{}}}}"
                )
            )
            for t in all_types:
                if re.search(r"inter.*laminar.*stress", t, re.IGNORECASE):
                    actual_type = t
                    break

            all_comps = _split_tcl_list(
                eval_tcl(
                    f'if {{![catch {{res GetDataComponentList {first_sub} "{actual_type}"}} _c]}} {{return $_c}} else {{return {{}}}}'
                )
            )
            for c in all_comps:
                if "yz" in c.lower():
                    comp_yz = c
                if "xz" in c.lower():
                    comp_xz = c

            rows = []
            g_max_xz = g_min_xz = g_max_yz = g_min_yz = None

            for sub in subcase_list:
                eval_tcl(f"res SetCurrentSubcase {sub}")
                sub_name = eval_tcl(f"res GetSubcaseLabel {sub}").replace(",", "-")

                max_yz = min_yz = max_xz = min_xz = None
                if num_elems > 0:
                    res_yz = _split_tcl_list(
                        eval_tcl(
                            f"""
                            cont SetEnableState false
                            cont SetDataType "{actual_type}"
                            if {{![catch {{cont SetDataComponent "{comp_yz}"}}]}} {{
                                catch {{cont SetShellLayer "Max"}}
                                cont SetEnableState true
                                cli Draw
                                ::pyhelper_QueryMinMax qry
                            }} else {{
                                list "" ""
                            }}
                            """
                        )
                    )
                    if len(res_yz) >= 2:
                        max_yz = _to_float(res_yz[0])
                        min_yz = _to_float(res_yz[1])

                    res_xz = _split_tcl_list(
                        eval_tcl(
                            f"""
                            cont SetEnableState false
                            cont SetDataType "{actual_type}"
                            if {{![catch {{cont SetDataComponent "{comp_xz}"}}]}} {{
                                catch {{cont SetShellLayer "Max"}}
                                cont SetEnableState true
                                cli Draw
                                ::pyhelper_QueryMinMax qry
                            }} else {{
                                list "" ""
                            }}
                            """
                        )
                    )
                    if len(res_xz) >= 2:
                        max_xz = _to_float(res_xz[0])
                        min_xz = _to_float(res_xz[1])

                if max_xz is not None and (g_max_xz is None or max_xz > g_max_xz):
                    g_max_xz = max_xz
                if min_xz is not None and (g_min_xz is None or min_xz < g_min_xz):
                    g_min_xz = min_xz
                if max_yz is not None and (g_max_yz is None or max_yz > g_max_yz):
                    g_max_yz = max_yz
                if min_yz is not None and (g_min_yz is None or min_yz < g_min_yz):
                    g_min_yz = min_yz

                rows.append(
                    [
                        sub,
                        sub_name,
                        max_xz,
                        min_xz,
                        get_extreme_float(max_xz, min_xz),
                        max_yz,
                        min_yz,
                        get_extreme_float(max_yz, min_yz),
                    ]
                )

            summary_row = [
                "MAX/MIN",
                "ENVELOPE (WORST CASE)",
                g_max_xz,
                g_min_xz,
                get_extreme_float(g_max_xz, g_min_xz),
                g_max_yz,
                g_min_yz,
                get_extreme_float(g_max_yz, g_min_yz),
            ]

            clusters = [
                {
                    "title": "IDENTIFICATION",
                    "color": "D9E1F2",
                    "sub_color": "EDF2F8",
                    "columns": ["Loadcase_ID", "Loadcase_Name"],
                },
                {
                    "title": "INTER-LAMINAR SHEAR XZ",
                    "color": "DDEBF7",
                    "sub_color": "EBF3FA",
                    "columns": ["Max_Shear_XZ", "Min_Shear_XZ", "Extreme_XZ"],
                },
                {
                    "title": "INTER-LAMINAR SHEAR YZ",
                    "color": "FCE4D6",
                    "sub_color": "FDF0E6",
                    "columns": ["Max_Shear_YZ", "Min_Shear_YZ", "Extreme_YZ"],
                },
            ]
            saved = export_clustered_report(
                filepath,
                "InterLaminar_Stress",
                clusters,
                rows,
                summary_row=summary_row,
            )
            messagebox.showinfo(
                "Export Complete",
                "Inter Laminar Stress Export complete!\nSaved to:\n"
                + "\n".join(saved),
                parent=self.win,
            )
        finally:
            self._hv_cleanup_handles(has_selection_set=True)

    # --- HV Part 5: Export Extreme Layer Normal/Shear Stresses (XX, YY, ZZ, XY, ZX, YZ) ---
    # Organized into 4 self-contained criterion clusters:
    #   1. MAX NORMAL (TENSION: Ftu | Max_XX | RF | Max_YY | RF | Max_ZZ | RF)
    #   2. MAX SHEAR  (SHEAR:   Fsu | Max_XY | RF | Max_ZX | RF | Max_YZ | RF)
    #   3. MIN NORMAL (COMPR.:  Fcu | Min_XX | RF | Min_YY | RF | Min_ZZ | RF)
    #   4. MIN SHEAR  (SHEAR:   Fsu | Min_XY | RF | Min_ZX | RF | Min_YZ | RF)
    def export_stress_extreme_layer(self):
        ftu_str = self.ftu.get()
        fcu_str = self.fcu.get()
        fsu_str = self.fsu.get()
        ftu_val = _to_float(ftu_str)
        fcu_val = _to_float(fcu_str)
        fsu_val = _to_float(fsu_str)
        # Note: in original TCLhelper5, shear RF used Ftu if desired, or Fsu.
        # Here we display both Ftu (for normal) and Fsu/Ftu clearly in each cluster.
        shear_allow = fsu_val if fsu_val is not None else ftu_val

        try:
            num_elems = self._hv_init_handles(include_query=True)
            if num_elems == 0:
                messagebox.showwarning(
                    "Warning",
                    "Warning: Your selection resulted in 0 elements.",
                    parent=self.win,
                )
                return

            filepath = self._ask_save_xlsx(
                "Save Extreme Normal and Shear Stresses",
                default_name=f"Stress_Extreme_Layer_{self._clean_set_name()}.xlsx",
            )
            if not filepath:
                return

            subcase_list = _split_tcl_list(eval_tcl('res GetSubcaseList "Base"'))
            if not subcase_list:
                return
            first_sub = subcase_list[0]

            actual_type = "Stresses"
            comp_list = _split_tcl_list(
                eval_tcl(f'res GetDataComponentList {first_sub} "{actual_type}"')
            )
            if not comp_list:
                actual_type = "Stress"
                comp_list = _split_tcl_list(
                    eval_tcl(f'res GetDataComponentList {first_sub} "{actual_type}"')
                )

            c_xx, c_yy, c_zz = "XX", "YY", "ZZ"
            c_xy, c_zx, c_yz = "XY", "ZX", "YZ"

            for c in comp_list:
                cl = c.lower()
                if "max" in cl:
                    continue
                if "xx" in cl:
                    c_xx = c
                if "yy" in cl:
                    c_yy = c
                if "zz" in cl:
                    c_zz = c
                if "xy" in cl:
                    c_xy = c
                if "zx" in cl:
                    c_zx = c
                if "yz" in cl:
                    c_yz = c

            components = [c_xx, c_yy, c_zz, c_xy, c_zx, c_yz]
            rows = []
            g_max = [None] * 6
            g_min = [None] * 6

            for sub in subcase_list:
                eval_tcl(f"res SetCurrentSubcase {sub}")
                sub_name = eval_tcl(f"res GetSubcaseLabel {sub}").replace(",", "-")

                l_max = [None] * 6
                l_min = [None] * 6

                if num_elems > 0:
                    for i, comp_name in enumerate(components):
                        res_pair = _split_tcl_list(
                            eval_tcl(
                                f"""
                                cont SetEnableState false
                                cont SetDataType "{actual_type}"
                                if {{![catch {{cont SetDataComponent "{comp_name}"}}]}} {{
                                    catch {{cont SetShellLayer "Extreme"}}
                                    catch {{cont SetResultSystem -1}}
                                    cont SetEnableState true
                                    cli Draw
                                    ::pyhelper_QueryMinMax qry
                                }} else {{
                                    list "" ""
                                }}
                                """
                            )
                        )
                        if len(res_pair) >= 2:
                            l_max[i] = _to_float(res_pair[0])
                            l_min[i] = _to_float(res_pair[1])

                for i in range(6):
                    if l_max[i] is not None and (
                        g_max[i] is None or l_max[i] > g_max[i]
                    ):
                        g_max[i] = l_max[i]
                    if l_min[i] is not None and (
                        g_min[i] is None or l_min[i] < g_min[i]
                    ):
                        g_min[i] = l_min[i]

                row = [
                    sub,
                    sub_name,
                    # Cluster 1: Max Normal (Ftu, Max_XX, RF, Max_YY, RF, Max_ZZ, RF)
                    ftu_val if ftu_val is not None else ftu_str,
                    l_max[0],
                    safe_rf(ftu_val, l_max[0]),
                    l_max[1],
                    safe_rf(ftu_val, l_max[1]),
                    l_max[2],
                    safe_rf(ftu_val, l_max[2]),
                    # Cluster 2: Max Shear (Fsu, Max_XY, RF, Max_ZX, RF, Max_YZ, RF)
                    shear_allow if shear_allow is not None else fsu_str,
                    l_max[3],
                    safe_rf(shear_allow, l_max[3]),
                    l_max[4],
                    safe_rf(shear_allow, l_max[4]),
                    l_max[5],
                    safe_rf(shear_allow, l_max[5]),
                    # Cluster 3: Min Normal (Fcu, Min_XX, RF, Min_YY, RF, Min_ZZ, RF)
                    fcu_val if fcu_val is not None else fcu_str,
                    l_min[0],
                    safe_rf(fcu_val, l_min[0]),
                    l_min[1],
                    safe_rf(fcu_val, l_min[1]),
                    l_min[2],
                    safe_rf(fcu_val, l_min[2]),
                    # Cluster 4: Min Shear (Fsu, Min_XY, RF, Min_ZX, RF, Min_YZ, RF)
                    shear_allow if shear_allow is not None else fsu_str,
                    l_min[3],
                    safe_rf(shear_allow, l_min[3]),
                    l_min[4],
                    safe_rf(shear_allow, l_min[4]),
                    l_min[5],
                    safe_rf(shear_allow, l_min[5]),
                ]
                rows.append(row)

            summary_row = [
                "MAX/MIN",
                "ENVELOPE (WORST CASE)",
                ftu_val if ftu_val is not None else ftu_str,
                g_max[0],
                safe_rf(ftu_val, g_max[0]),
                g_max[1],
                safe_rf(ftu_val, g_max[1]),
                g_max[2],
                safe_rf(ftu_val, g_max[2]),
                shear_allow if shear_allow is not None else fsu_str,
                g_max[3],
                safe_rf(shear_allow, g_max[3]),
                g_max[4],
                safe_rf(shear_allow, g_max[4]),
                g_max[5],
                safe_rf(shear_allow, g_max[5]),
                fcu_val if fcu_val is not None else fcu_str,
                g_min[0],
                safe_rf(fcu_val, g_min[0]),
                g_min[1],
                safe_rf(fcu_val, g_min[1]),
                g_min[2],
                safe_rf(fcu_val, g_min[2]),
                shear_allow if shear_allow is not None else fsu_str,
                g_min[3],
                safe_rf(shear_allow, g_min[3]),
                g_min[4],
                safe_rf(shear_allow, g_min[4]),
                g_min[5],
                safe_rf(shear_allow, g_min[5]),
            ]

            clusters = [
                {
                    "title": "IDENTIFICATION",
                    "color": "D9E1F2",
                    "sub_color": "EDF2F8",
                    "columns": ["Loadcase_ID", "Loadcase_Name"],
                },
                {
                    "title": "MAX NORMAL STRESSES (TENSION)",
                    "color": "DDEBF7",
                    "sub_color": "EBF3FA",
                    "columns": [
                        "Ftu",
                        "Max_XX",
                        "RF_Max_XX",
                        "Max_YY",
                        "RF_Max_YY",
                        "Max_ZZ",
                        "RF_Max_ZZ",
                    ],
                },
                {
                    "title": "MAX SHEAR STRESSES (XY, ZX, YZ)",
                    "color": "E2EFDA",
                    "sub_color": "EFF7EB",
                    "columns": [
                        "Fsu",
                        "Max_XY",
                        "RF_Max_XY",
                        "Max_ZX",
                        "RF_Max_ZX",
                        "Max_YZ",
                        "RF_Max_YZ",
                    ],
                },
                {
                    "title": "MIN NORMAL STRESSES (COMPRESSION)",
                    "color": "FFF2CC",
                    "sub_color": "FFF8E5",
                    "columns": [
                        "Fcu",
                        "Min_XX",
                        "RF_Min_XX",
                        "Min_YY",
                        "RF_Min_YY",
                        "Min_ZZ",
                        "RF_Min_ZZ",
                    ],
                },
                {
                    "title": "MIN SHEAR STRESSES (XY, ZX, YZ)",
                    "color": "FCE4D6",
                    "sub_color": "FDF0E6",
                    "columns": [
                        "Fsu",
                        "Min_XY",
                        "RF_Min_XY",
                        "Min_ZX",
                        "RF_Min_ZX",
                        "Min_YZ",
                        "RF_Min_YZ",
                    ],
                },
            ]

            rf_indices = {4, 6, 8, 11, 13, 15, 18, 20, 22, 25, 27, 29}
            saved = export_clustered_report(
                filepath,
                "Layer_Stresses_All",
                clusters,
                rows,
                rf_col_indices=rf_indices,
                summary_row=summary_row,
            )
            messagebox.showinfo(
                "Export Complete",
                "Normal & Shear Stresses Export complete!\nSaved to:\n"
                + "\n".join(saved),
                parent=self.win,
            )
        finally:
            self._hv_cleanup_handles(has_selection_set=True)

    # --- HV Part 6: Export Extreme Layer Stresses (ZX & YZ Only) ---
    # Organized into 2 self-contained criterion clusters:
    #   1. SHEAR ZX CRITERION: [Fzx | Max_ZX | RF_Max_ZX | Min_ZX | RF_Min_ZX]
    #   2. SHEAR YZ CRITERION: [Fyz | Max_YZ | RF_Max_YZ | Min_YZ | RF_Min_YZ]
    def export_stress_extreme_layer_zxyz(self):
        fzx_str = self.fzx.get()
        fyz_str = self.fyz.get()
        fzx_val = _to_float(fzx_str)
        fyz_val = _to_float(fyz_str)

        try:
            num_elems = self._hv_init_handles(include_query=True)
            if num_elems == 0:
                messagebox.showwarning(
                    "Warning",
                    "Warning: Your selection resulted in 0 elements.",
                    parent=self.win,
                )
                return

            filepath = self._ask_save_xlsx(
                "Save Extreme Normal and Shear Stresses (ZX & YZ)",
                default_name=f"Stress_ZX_YZ_{self._clean_set_name()}.xlsx",
            )
            if not filepath:
                return

            subcase_list = _split_tcl_list(eval_tcl('res GetSubcaseList "Base"'))
            if not subcase_list:
                return
            first_sub = subcase_list[0]

            actual_type = "Stresses"
            comp_list = _split_tcl_list(
                eval_tcl(f'res GetDataComponentList {first_sub} "{actual_type}"')
            )
            if not comp_list:
                actual_type = "Stress"
                comp_list = _split_tcl_list(
                    eval_tcl(f'res GetDataComponentList {first_sub} "{actual_type}"')
                )

            c_zx, c_yz = "ZX", "YZ"
            for c in comp_list:
                cl = c.lower()
                if "max" in cl:
                    continue
                if "zx" in cl:
                    c_zx = c
                if "yz" in cl:
                    c_yz = c

            components = [c_zx, c_yz]
            rows = []
            g_max = [None, None]
            g_min = [None, None]

            for sub in subcase_list:
                eval_tcl(f"res SetCurrentSubcase {sub}")
                sub_name = eval_tcl(f"res GetSubcaseLabel {sub}").replace(",", "-")

                l_max = [None, None]
                l_min = [None, None]

                if num_elems > 0:
                    for i, comp_name in enumerate(components):
                        res_pair = _split_tcl_list(
                            eval_tcl(
                                f"""
                                cont SetEnableState false
                                cont SetDataType "{actual_type}"
                                if {{![catch {{cont SetDataComponent "{comp_name}"}}]}} {{
                                    catch {{cont SetShellLayer "Extreme"}}
                                    catch {{cont SetResultSystem -1}}
                                    cont SetEnableState true
                                    cli Draw
                                    ::pyhelper_QueryMinMax qry
                                }} else {{
                                    list "" ""
                                }}
                                """
                            )
                        )
                        if len(res_pair) >= 2:
                            l_max[i] = _to_float(res_pair[0])
                            l_min[i] = _to_float(res_pair[1])

                for i in range(2):
                    if l_max[i] is not None and (
                        g_max[i] is None or l_max[i] > g_max[i]
                    ):
                        g_max[i] = l_max[i]
                    if l_min[i] is not None and (
                        g_min[i] is None or l_min[i] < g_min[i]
                    ):
                        g_min[i] = l_min[i]

                rows.append(
                    [
                        sub,
                        sub_name,
                        # Cluster 1: Shear ZX [Fzx, Max_ZX, RF_Max_ZX, Min_ZX, RF_Min_ZX]
                        fzx_val if fzx_val is not None else fzx_str,
                        l_max[0],
                        safe_rf(fzx_val, l_max[0]),
                        l_min[0],
                        safe_rf(fzx_val, l_min[0]),
                        # Cluster 2: Shear YZ [Fyz, Max_YZ, RF_Max_YZ, Min_YZ, RF_Min_YZ]
                        fyz_val if fyz_val is not None else fyz_str,
                        l_max[1],
                        safe_rf(fyz_val, l_max[1]),
                        l_min[1],
                        safe_rf(fyz_val, l_min[1]),
                    ]
                )

            summary_row = [
                "MAX/MIN",
                "ENVELOPE (WORST CASE)",
                fzx_val if fzx_val is not None else fzx_str,
                g_max[0],
                safe_rf(fzx_val, g_max[0]),
                g_min[0],
                safe_rf(fzx_val, g_min[0]),
                fyz_val if fyz_val is not None else fyz_str,
                g_max[1],
                safe_rf(fyz_val, g_max[1]),
                g_min[1],
                safe_rf(fyz_val, g_min[1]),
            ]

            clusters = [
                {
                    "title": "IDENTIFICATION",
                    "color": "D9E1F2",
                    "sub_color": "EDF2F8",
                    "columns": ["Loadcase_ID", "Loadcase_Name"],
                },
                {
                    "title": "OUT-OF-PLANE SHEAR ZX CRITERION",
                    "color": "DDEBF7",
                    "sub_color": "EBF3FA",
                    "columns": ["Fzx", "Max_ZX", "RF_Max_ZX", "Min_ZX", "RF_Min_ZX"],
                },
                {
                    "title": "OUT-OF-PLANE SHEAR YZ CRITERION",
                    "color": "FCE4D6",
                    "sub_color": "FDF0E6",
                    "columns": ["Fyz", "Max_YZ", "RF_Max_YZ", "Min_YZ", "RF_Min_YZ"],
                },
            ]

            saved = export_clustered_report(
                filepath,
                "Stresses_ZX_YZ",
                clusters,
                rows,
                rf_col_indices={4, 6, 9, 11},
                summary_row=summary_row,
            )
            messagebox.showinfo(
                "Export Complete",
                "ZX & YZ Stresses Export complete!\nSaved to:\n" + "\n".join(saved),
                parent=self.win,
            )
        finally:
            self._hv_cleanup_handles(has_selection_set=True)

    # --- HV Part 7: Export 1D Element Fastener Forces (Clustered Tension & Shear) ---
    def export_fastener_forces(self):
        ft_str = self.ft.get()
        fs_str = self.fs.get()
        ft_val = _to_float(ft_str)
        fs_val = _to_float(fs_str)

        try:
            num_elems = self._hv_init_handles(include_query=True)
            if num_elems == 0:
                messagebox.showwarning(
                    "Warning",
                    "Warning: Your selection resulted in 0 elements.",
                    parent=self.win,
                )
                return

            filepath = self._ask_save_xlsx(
                "Save 1D Element Forces",
                default_name=f"Fastener_Forces_{self._clean_set_name()}.xlsx",
            )
            if not filepath:
                return

            subcase_list = _split_tcl_list(eval_tcl('res GetSubcaseList "Base"'))
            if not subcase_list:
                return
            first_sub = subcase_list[0]

            actual_type = "1D Element Forces"
            all_types = _split_tcl_list(
                eval_tcl(
                    f"if {{![catch {{res GetDataTypeList {first_sub}}} _t]}} {{return $_t}} else {{return {{}}}}"
                )
            )
            for t in all_types:
                if re.search(r"1d.*force", t, re.IGNORECASE):
                    actual_type = t
                    break

            comp_x, comp_y, comp_z = "X", "Y", "Z"
            all_comps = _split_tcl_list(
                eval_tcl(
                    f'if {{![catch {{res GetDataComponentList {first_sub} "{actual_type}"}} _c]}} {{return $_c}} else {{return {{}}}}'
                )
            )
            for c in all_comps:
                cl = c.lower()
                if cl in ("x", "x-axis"):
                    comp_x = c
                elif cl in ("y", "y-axis"):
                    comp_y = c
                elif cl in ("z", "z-axis"):
                    comp_z = c

            rows = []
            max_abs_fx = None
            max_abs_fy = None
            max_abs_fz = None
            max_shear = None
            worst_eid = "-"

            for sub in subcase_list:
                eval_tcl(f"res SetCurrentSubcase {sub}")
                sub_name = eval_tcl(f"res GetSubcaseLabel {sub}").replace(",", "-")

                val_x = {}
                val_y = {}
                val_z = {}

                if num_elems > 0:
                    for comp_name, target_dict in (
                        (comp_x, val_x),
                        (comp_y, val_y),
                        (comp_z, val_z),
                    ):
                        raw_pairs = eval_tcl(
                            f"""
                            cont SetEnableState false
                            cont SetDataType "{actual_type}"
                            if {{![catch {{cont SetDataComponent "{comp_name}"}}]}} {{
                                cont SetEnableState true
                                cli Draw
                                ::pyhelper_QueryElementAbsExtreme qry
                            }} else {{
                                return ""
                            }}
                            """
                        )
                        for pair in raw_pairs.split():
                            if ":" in pair:
                                eid_s, val_s = pair.split(":", 1)
                                f_v = _to_float(val_s)
                                if f_v is not None:
                                    target_dict[int(eid_s)] = f_v

                for eid in sorted(val_x.keys()):
                    fx = _clean_zero(val_x[eid], 2)
                    fy = _clean_zero(val_y.get(eid, 0.0), 2)
                    fz = _clean_zero(val_z.get(eid, 0.0), 2)

                    rf_tension = safe_rf(ft_val, fx)
                    shear = _clean_zero(math.hypot(fy, fz), 2)
                    rf_shear = safe_rf(fs_val, shear)

                    if max_abs_fx is None or abs(fx) > abs(max_abs_fx):
                        max_abs_fx = fx
                    if max_abs_fy is None or abs(fy) > abs(max_abs_fy):
                        max_abs_fy = fy
                    if max_abs_fz is None or abs(fz) > abs(max_abs_fz):
                        max_abs_fz = fz
                    if max_shear is None or shear > max_shear:
                        max_shear = shear
                        worst_eid = eid

                    rows.append(
                        [
                            sub,
                            sub_name,
                            eid,
                            ft_val if ft_val is not None else ft_str,
                            fx,
                            rf_tension,
                            fs_val if fs_val is not None else fs_str,
                            fy,
                            fz,
                            shear,
                            rf_shear,
                        ]
                    )

            summary_row = None
            if rows:
                summary_row = [
                    "MAX",
                    "ENVELOPE (WORST CASE)",
                    worst_eid,
                    ft_val if ft_val is not None else ft_str,
                    max_abs_fx,
                    safe_rf(ft_val, max_abs_fx),
                    fs_val if fs_val is not None else fs_str,
                    max_abs_fy,
                    max_abs_fz,
                    max_shear,
                    safe_rf(fs_val, max_shear),
                ]

            clusters = [
                {
                    "title": "IDENTIFICATION",
                    "color": "D9E1F2",
                    "sub_color": "EDF2F8",
                    "columns": ["Loadcase_ID", "Loadcase_Name", "Element_ID"],
                },
                {
                    "title": "TENSION ANALYSIS (AXIAL X)",
                    "color": "DDEBF7",
                    "sub_color": "EBF3FA",
                    "columns": ["Ft", "Force_X", "RF_Tension"],
                },
                {
                    "title": "SHEAR ANALYSIS (PLANE YZ)",
                    "color": "FCE4D6",
                    "sub_color": "FDF0E6",
                    "columns": ["Fs", "Force_Y", "Force_Z", "Shear_Force", "RF_Shear"],
                },
            ]

            saved = export_clustered_report(
                filepath,
                "1D_Fastener_Forces",
                clusters,
                rows,
                rf_col_indices={5, 10},
                summary_row=summary_row,
            )

            messagebox.showinfo(
                "Export Complete",
                "Fastener Forces Export complete!\nSaved to:\n" + "\n".join(saved),
                parent=self.win,
            )
        finally:
            self._hv_cleanup_handles(has_selection_set=True)


_APP_INSTANCE = None


def main():
    global _APP_INSTANCE
    if _APP_INSTANCE is not None:
        try:
            _APP_INSTANCE.win.destroy()
        except Exception:
            pass
    _APP_INSTANCE = TCLHelperApp()
    if not _INSIDE_HYPERWORKS:
        _APP_INSTANCE.win.mainloop()
    return _APP_INSTANCE


if __name__ == "__main__":
    main()
