import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# --- Global State ---
selected_files = []

# ─────────────────────────────────────────────
# Extraction helpers
# ─────────────────────────────────────────────

def extract_grid_points(f06_path):
    grid_points = []
    pattern = r"GRID POINT\s+(\d+)\s+COMPONENT\s+1 ILLEGALLY DEFINED"
    with open(f06_path, "r") as file:
        for line in file:
            match = re.search(pattern, line)
            if match:
                grid_points.append(match.group(1))
    return grid_points

def extract_elements_with_curvature_issue(f06_path):
    elements = []
    pattern = r"ELEMENT\s+(\d+)\s+HAS TOO MUCH CURVATURE"
    with open(f06_path, "r") as file:
        for line in file:
            match = re.search(pattern, line)
            if match:
                elements.append(match.group(1))
    return elements

def extract_user_fatal_2101_grid_points(f06_path):
    grid_points = []
    pattern_error_start = r"\*\*\* USER FATAL MESSAGE 2101"
    pattern_grid_point_id = r"GRID POINT\s+(\d+)\s+COMPONENT"
    with open(f06_path, "r") as file:
        lines = file.readlines()
    for i, line in enumerate(lines):
        if re.search(pattern_error_start, line):
            match = re.search(pattern_grid_point_id, line)
            if match:
                grid_points.append(match.group(1))
            else:
                for j in range(i+1, min(i+5, len(lines))):
                    match = re.search(pattern_grid_point_id, lines[j])
                    if match:
                        grid_points.append(match.group(1))
                        break
    return grid_points

def extract_user_fatal_7546_ids(f06_path):
    element_ids = []
    pattern_error_start = r"\*\*\* USER FATAL MESSAGE 7546"
    pattern_element_id = r"THE SHID1 \(OR SHID2\)=\s+(\d+)"
    with open(f06_path, "r") as file:
        lines = file.readlines()
    for i, line in enumerate(lines):
        if re.search(pattern_error_start, line):
            for j in range(i+1, min(i+5, len(lines))):
                match = re.search(pattern_element_id, lines[j])
                if match:
                    element_ids.append(match.group(1))
                    break
    return element_ids

def extract_user_fatal_7549_ids(f06_path):
    element_ids = []
    pattern_error_start = r"\*\*\* USER FATAL MESSAGE 7549"
    pattern_element_id = r"CFAST ELEMENT ID=(\d+)"
    with open(f06_path, "r") as file:
        lines = file.readlines()
    for i, line in enumerate(lines):
        if re.search(pattern_error_start, line):
            match = re.search(pattern_element_id, line)
            if match:
                element_ids.append(match.group(1))
            else:
                for j in range(i+1, min(i+5, len(lines))):
                    match = re.search(pattern_element_id, lines[j])
                    if match:
                        element_ids.append(match.group(1))
                        break
    return element_ids

def extract_system_not_definable_ids(f06_path):
    ids = []
    pattern = r"Error:\s+System\s+(\d+)\s+is not definable"
    try:
        with open(f06_path, "r") as file:
            for line in file:
                match = re.search(pattern, line)
                if match:
                    ids.append(match.group(1))
    except Exception as e:
        print(f"Error reading file for System IDs: {e}")
    return ids

def has_fatal_in_file(f06_path):
    """Return True if any USER FATAL MESSAGE is found in the file."""
    pattern = r"\*\*\* USER FATAL MESSAGE"
    try:
        with open(f06_path, "r") as file:
            for line in file:
                if re.search(pattern, line):
                    return True
    except Exception:
        pass
    return False

# ─────────────────────────────────────────────
# Maximum Displacements extraction
# ─────────────────────────────────────────────

def extract_max_displacements(f06_path):
    """
    Parse all MAXIMUM DISPLACEMENTS blocks from an F06 file.
    Returns a list of dicts with keys: subcase, T1, T2, T3, R1, R2, R3
    Multiple occurrences of the same subcase are all kept (one row per block line).
    """
    results = []
    # Pattern for the section header
    header_pattern = re.compile(r'MAXIMUM\s+DISPLACEMENTS', re.IGNORECASE)
    # Pattern for data lines: starts with optional spaces then '0' then subcase id
    # e.g.  "0     1001   2.7940121E+01  1.2191954E+00 ..."
    data_pattern = re.compile(
        r'^\s*0\s+(\d+)\s+'
        r'([+-]?\d+\.\d+[Ee][+-]?\d+)\s+'
        r'([+-]?\d+\.\d+[Ee][+-]?\d+)\s+'
        r'([+-]?\d+\.\d+[Ee][+-]?\d+)\s+'
        r'([+-]?\d+\.\d+[Ee][+-]?\d+)\s+'
        r'([+-]?\d+\.\d+[Ee][+-]?\d+)\s+'
        r'([+-]?\d+\.\d+[Ee][+-]?\d+)'
    )

    with open(f06_path, "r", errors="replace") as f:
        lines = f.readlines()

    in_section = False
    for i, line in enumerate(lines):
        if header_pattern.search(line):
            in_section = True
            continue
        if in_section:
            # blank line or new page (form feed) ends section
            stripped = line.strip()
            if stripped == "" or stripped.startswith("1 "):
                in_section = False
                continue
            m = data_pattern.match(line)
            if m:
                results.append({
                    "subcase": int(m.group(1)),
                    "T1": float(m.group(2)),
                    "T2": float(m.group(3)),
                    "T3": float(m.group(4)),
                    "R1": float(m.group(5)),
                    "R2": float(m.group(6)),
                    "R3": float(m.group(7)),
                })
    return results

# ─────────────────────────────────────────────
# Force balance helpers
# ─────────────────────────────────────────────

def parse_totals_line(line_parts):
    values = []
    for v in line_parts:
        try:
            values.append(float(v))
        except ValueError:
            values.append(0.0)
    return values

def sum_totals(values):
    return sum(values)

def compare_totals(v1, v2):
    return abs(sum_totals(v1) - sum_totals(v2)) < 1e-12

def process_file(filename):
    with open(filename, "r") as f:
        lines = f.readlines()

    section = None
    subcase = None
    oload_totals = {}
    spcforce_totals = {}
    mpcforce_totals = {}

    def parse_val(v):
        """Convert value string to float; treat '----' dashes as 0.0."""
        if re.match(r'^-{2,}$', v.strip()):
            return 0.0
        try:
            return float(v)
        except ValueError:
            return 0.0

    i = 0
    while i < len(lines):
        line = lines[i]
        strip = line.strip()

        if re.search(r'OLOAD\s+RESULTANT', line):
            section = "OLOAD"
            i += 1; continue
        elif re.search(r'SPCFORCE\s+RESULTANT', line):
            section = "SPCFORCE"
            i += 1; continue
        elif re.search(r'MPCFORCE\s+RESULTANT', line):
            section = "MPCFORCE"
            i += 1; continue

        if section and strip == "":
            section = None
            i += 1; continue

        subcase_match = re.match(r'^0\s+(\d+)\s+', line)
        if subcase_match:
            subcase = int(subcase_match.group(1))
            i += 1; continue

        if section and re.match(r'^\s+TOTALS\s+', line):
            parts = strip.split()
            total_strs = parts[1:]

            j = i + 1
            while len(total_strs) < 6 and j < len(lines):
                next_strip = lines[j].strip()
                if (next_strip == "" or
                    re.match(r'^\s+TOTALS\s+', lines[j]) or
                    re.match(r'^0\s+\d+\s+', lines[j]) or
                    re.search(r'(OLOAD|SPCFORCE|MPCFORCE)\s+RESULTANT', lines[j])):
                    break
                total_strs.extend(next_strip.split())
                j += 1

            totals_values = [parse_val(v) for v in total_strs[:6]]

            if subcase is not None:
                if section == "OLOAD":
                    oload_totals[subcase] = totals_values
                elif section == "SPCFORCE":
                    spcforce_totals[subcase] = totals_values
                elif section == "MPCFORCE":
                    mpcforce_totals[subcase] = totals_values
            i = j; continue

        i += 1

    results = []
    all_subcases = (set(oload_totals.keys()) |
                    set(spcforce_totals.keys()) |
                    set(mpcforce_totals.keys()))
    for sc in sorted(all_subcases):
        ovals = oload_totals.get(sc)
        svals = spcforce_totals.get(sc)
        mvals = mpcforce_totals.get(sc)
        if not ovals or not svals or not mvals:
            results.append((sc, "Missing data", "black"))
        elif compare_totals(ovals, svals) and compare_totals(ovals, mvals):
            results.append((sc, "OK", "green"))
        else:
            results.append((sc, "not OK", "red"))

    return results, oload_totals, spcforce_totals, mpcforce_totals

# ─────────────────────────────────────────────
# Display helpers
# ─────────────────────────────────────────────

def display_results(results, oload_totals, spcforce_totals, mpcforce_totals):
    col_width = 18
    sep = "|"

    def fmt_val(v):
        if v == 0.0:
            return "0"
        abs_v = abs(v)
        if abs_v >= 1:
            if abs_v >= 1000:
                return f"{v:,.1f}"
            else:
                return f"{v:.2f}"
        else:
            return f"{v:.6f}"

    def draw_separator(w):
        w.insert(tk.END, "+" + "+".join(["-" * col_width] * 8) + "+\n")

    text_area_loads.config(state="normal")
    text_area_loads.delete("1.0", tk.END)

    for sc, status_text, color_tag in results:
        ovals = oload_totals.get(sc)
        svals = spcforce_totals.get(sc)
        mvals = mpcforce_totals.get(sc)

        text_area_loads.insert(tk.END, f"Subcase {sc}:\n")
        draw_separator(text_area_loads)

        header_type = f"{'Type':^{col_width}}"
        header_cols = "".join(f"{c:^{col_width}}" for c in ["T1","T2","T3","R1","R2","R3"])
        text_area_loads.insert(tk.END, sep + header_type + header_cols + sep + "\n")
        draw_separator(text_area_loads)

        for label, vals in [("OLOAD", ovals), ("SPCFORCE", svals), ("MPCFORCE", mvals)]:
            if vals:
                row = f"{label:^{col_width}}" + "".join(f"{fmt_val(v):>{col_width}}" for v in vals)
            else:
                row = f"{label:^{col_width}}" + f"{'Missing':^{col_width * 6}}"
            text_area_loads.insert(tk.END, sep + row + sep + "\n")

        text_area_loads.insert(tk.END, sep)
        text_area_loads.insert(tk.END, f"{'Status':^{col_width}}")

        pad_l = (col_width - len(status_text)) // 2
        pad_r = col_width - pad_l - len(status_text)
        text_area_loads.insert(tk.END, " " * pad_l)
        idx_start = text_area_loads.index(tk.END)
        text_area_loads.insert(tk.END, status_text)
        idx_end = text_area_loads.index(tk.END)
        text_area_loads.tag_add(color_tag, idx_start, idx_end)
        text_area_loads.insert(tk.END, " " * pad_r + sep + "\n")
        draw_separator(text_area_loads)
        text_area_loads.insert(tk.END, "\n")

    text_area_loads.config(state="disabled")


def display_max_displacements(disp_rows):
    """Render the displacement data in the dedicated tab Treeview."""
    # Clear existing rows
    for row in tree_disp.get_children():
        tree_disp.delete(row)

    for r in disp_rows:
        tree_disp.insert("", tk.END, values=(
            r["subcase"],
            f"{r['T1']:.2f}",
            f"{r['T2']:.2f}",
            f"{r['T3']:.2f}",
            f"{r['R1']:.2f}",
            f"{r['R2']:.2f}",
            f"{r['R3']:.2f}",
        ))

    lbl_disp_count.config(text=f"{len(disp_rows)} row(s) found")


def export_displacements_csv():
    """Export displacements from ALL loaded F06 files to one CSV."""
    if not selected_files:
        messagebox.showinfo("Export", "No F06 files loaded.")
        return

    save_path = filedialog.asksaveasfilename(
        title="Save all displacements as CSV",
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not save_path:
        return

    headers = ["filename", "subcase", "T1", "T2", "T3", "R1", "R2", "R3"]
    total_rows = 0
    with open(save_path, "w", newline="") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)
        for fp in selected_files:
            fname = fp.replace("\\", "/").split("/")[-1]
            try:
                rows = extract_max_displacements(fp)
                for r in rows:
                    writer.writerow([
                        fname,
                        r["subcase"],
                        f"{r['T1']:.2f}",
                        f"{r['T2']:.2f}",
                        f"{r['T3']:.2f}",
                        f"{r['R1']:.2f}",
                        f"{r['R2']:.2f}",
                        f"{r['R3']:.2f}",
                    ])
                    total_rows += 1
            except Exception as e:
                writer.writerow([fname, "ERROR", str(e), "", "", "", "", ""])

    messagebox.showinfo("Export", f"Exported {total_rows} row(s) from {len(selected_files)} file(s) to:\n{save_path}")


# ─────────────────────────────────────────────
# Update all tabs
# ─────────────────────────────────────────────

def update_all_results(filepath):
    """Populate all tabs with data extracted from filepath."""
    grid_points = extract_grid_points(filepath)
    text_area_grid.config(state="normal")
    text_area_grid.delete("1.0", tk.END)
    text_area_grid.insert(tk.END,
        f"Found {len(grid_points)} grid points:\n" + ", ".join(grid_points)
        if grid_points else "No grid points found.")
    text_area_grid.config(state="disabled")

    elements = extract_elements_with_curvature_issue(filepath)
    text_area_elements.config(state="normal")
    text_area_elements.delete("1.0", tk.END)
    text_area_elements.insert(tk.END,
        f"Found {len(elements)} elements with curvature issues:\n" + ", ".join(elements)
        if elements else "No elements with curvature issues found.")
    text_area_elements.config(state="disabled")

    ids7546 = extract_user_fatal_7546_ids(filepath)
    text_area_fatal_ids.config(state="normal")
    text_area_fatal_ids.delete("1.0", tk.END)
    text_area_fatal_ids.insert(tk.END,
        f"Found {len(ids7546)} element IDs in USER FATAL MESSAGE 7546:\n" + ", ".join(ids7546)
        if ids7546 else "No USER FATAL MESSAGE 7546 element IDs found.")
    text_area_fatal_ids.config(state="disabled")

    ids7549 = extract_user_fatal_7549_ids(filepath)
    text_area_fatal_7549_ids.config(state="normal")
    text_area_fatal_7549_ids.delete("1.0", tk.END)
    text_area_fatal_7549_ids.insert(tk.END,
        f"Found {len(ids7549)} element IDs in USER FATAL MESSAGE 7549:\n" + ", ".join(ids7549)
        if ids7549 else "No USER FATAL MESSAGE 7549 element IDs found.")
    text_area_fatal_7549_ids.config(state="disabled")

    ids2101 = extract_user_fatal_2101_grid_points(filepath)
    text_area_fatal_2101_grid_points.config(state="normal")
    text_area_fatal_2101_grid_points.delete("1.0", tk.END)
    text_area_fatal_2101_grid_points.insert(tk.END,
        f"Found {len(ids2101)} GRID POINT IDs in USER FATAL MESSAGE 2101:\n" + ", ".join(ids2101)
        if ids2101 else "No USER FATAL MESSAGE 2101 GRID POINT IDs found.")
    text_area_fatal_2101_grid_points.config(state="disabled")

    sys_ids = extract_system_not_definable_ids(filepath)
    text_area_system_definable.config(state="normal")
    text_area_system_definable.delete("1.0", tk.END)
    text_area_system_definable.insert(tk.END,
        f"Found {len(sys_ids)} System IDs not definable:\n" + ", ".join(sys_ids)
        if sys_ids else "No 'System not definable' errors found.")
    text_area_system_definable.config(state="disabled")

    results, oload_totals, spcforce_totals, mpcforce_totals = process_file(filepath)
    display_results(results, oload_totals, spcforce_totals, mpcforce_totals)

    disp_rows = extract_max_displacements(filepath)
    display_max_displacements(disp_rows)


# ─────────────────────────────────────────────
# File-list helpers
# ─────────────────────────────────────────────

def refresh_file_listbox():
    file_listbox.delete(0, tk.END)
    for fp in selected_files:
        fname = fp.replace("\\", "/").split("/")[-1]
        fatal = has_fatal_in_file(fp)
        label = f"⚠ {fname}" if fatal else fname
        file_listbox.insert(tk.END, label)
        if fatal:
            file_listbox.itemconfig(tk.END, {"fg": "red"})

def on_file_select(event):
    selection = file_listbox.curselection()
    if not selection:
        return
    idx = selection[0]
    filepath = selected_files[idx]
    fname = filepath.replace("\\", "/").split("/")[-1]
    label_selected_file.config(text=f"Showing: {fname}")
    update_all_results(filepath)

def select_f06_files():
    global selected_files
    filepaths = filedialog.askopenfilenames(
        title="Select one or more F06 files",
        filetypes=[("F06 files", "*.f06"), ("All files", "*.*")]
    )
    if not filepaths:
        return
    for fp in filepaths:
        if fp not in selected_files:
            selected_files.append(fp)
    refresh_file_listbox()
    first_new_idx = selected_files.index(filepaths[0])
    file_listbox.selection_clear(0, tk.END)
    file_listbox.selection_set(first_new_idx)
    file_listbox.see(first_new_idx)
    filepath = selected_files[first_new_idx]
    fname = filepath.replace("\\", "/").split("/")[-1]
    label_selected_file.config(text=f"Showing: {fname}")
    update_all_results(filepath)

def clear_file_list():
    global selected_files
    selected_files = []
    file_listbox.delete(0, tk.END)
    label_selected_file.config(text="No file selected")
    for ta in [text_area_grid, text_area_elements, text_area_loads,
               text_area_fatal_ids, text_area_fatal_7549_ids,
               text_area_fatal_2101_grid_points, text_area_system_definable]:
        ta.config(state="normal")
        ta.delete("1.0", tk.END)
        ta.config(state="disabled")
    for row in tree_disp.get_children():
        tree_disp.delete(row)
    lbl_disp_count.config(text="0 row(s) found")

# ─────────────────────────────────────────────
# GUI Layout
# ─────────────────────────────────────────────

root = tk.Tk()
root.title("F06 Reader – Multi-file")
root.geometry("1400x750")

frame_top = tk.Frame(root, pady=6)
frame_top.pack(fill="x")
tk.Button(frame_top, text="Add F06 File(s)", command=select_f06_files).pack(side="left", padx=8)
tk.Button(frame_top, text="Clear List", command=clear_file_list).pack(side="left", padx=4)
label_selected_file = tk.Label(frame_top, text="No file selected", anchor="w")
label_selected_file.pack(side="left", padx=10)

paned = tk.PanedWindow(root, orient=tk.HORIZONTAL, sashwidth=5, sashrelief=tk.RAISED)
paned.pack(fill="both", expand=True, padx=4, pady=4)

frame_left = tk.Frame(paned, width=230)
tk.Label(frame_left, text="F06 Files", font=("Consolas", 10, "bold")).pack(anchor="w", padx=4)
scrollbar_left = tk.Scrollbar(frame_left, orient=tk.VERTICAL)
file_listbox = tk.Listbox(frame_left, yscrollcommand=scrollbar_left.set,
                          font=("Consolas", 10), selectmode=tk.SINGLE, width=30)
scrollbar_left.config(command=file_listbox.yview)
scrollbar_left.pack(side=tk.RIGHT, fill=tk.Y)
file_listbox.pack(fill="both", expand=True, padx=2, pady=2)
file_listbox.bind("<<ListboxSelect>>", on_file_select)
paned.add(frame_left, minsize=180)

frame_right = tk.Frame(paned)
tab_control = ttk.Notebook(frame_right)
tab_control.pack(expand=True, fill="both")
paned.add(frame_right, minsize=800)

FONT = ("Consolas", 11)

def make_tab(title):
    tab = ttk.Frame(tab_control)
    tab_control.add(tab, text=title)
    area = scrolledtext.ScrolledText(tab, font=FONT, state="disabled")
    area.pack(fill="both", expand=True, padx=6, pady=6)
    return area

text_area_grid                  = make_tab("Grid Points")
text_area_elements              = make_tab("Curvature Issues")
text_area_loads                 = make_tab("OLOAD / SPC / MPC")
text_area_fatal_ids             = make_tab("ERROR 7546")
text_area_fatal_7549_ids        = make_tab("ERROR 7549")
text_area_fatal_2101_grid_points= make_tab("ERROR 2101")
text_area_system_definable      = make_tab("System Not Definable")

for tag, colour in [("green","green"), ("red","red"), ("black","black")]:
    text_area_loads.tag_configure(tag, foreground=colour)

# ── New tab: Maximum Displacements ──────────────────────────────────────────
tab_disp = ttk.Frame(tab_control)
tab_control.add(tab_disp, text="Max Displacements")

# toolbar inside the tab
frame_disp_top = tk.Frame(tab_disp)
frame_disp_top.pack(fill="x", padx=6, pady=(6, 2))
tk.Button(frame_disp_top, text="Export CSV",
          command=export_displacements_csv).pack(side="left", padx=4)
lbl_disp_count = tk.Label(frame_disp_top, text="0 row(s) found", anchor="w")
lbl_disp_count.pack(side="left", padx=10)

# Treeview with scrollbars
frame_disp_tree = tk.Frame(tab_disp)
frame_disp_tree.pack(fill="both", expand=True, padx=6, pady=(0, 6))

DISP_COLS = ("subcase", "T1", "T2", "T3", "R1", "R2", "R3")
tree_disp = ttk.Treeview(frame_disp_tree, columns=DISP_COLS, show="headings")

for col in DISP_COLS:
    tree_disp.heading(col, text=col)
    w = 80 if col != "subcase" else 90
    tree_disp.column(col, width=w, anchor="e" if col != "subcase" else "center")

vsb = ttk.Scrollbar(frame_disp_tree, orient="vertical",   command=tree_disp.yview)
hsb = ttk.Scrollbar(frame_disp_tree, orient="horizontal", command=tree_disp.xview)
tree_disp.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
vsb.pack(side=tk.RIGHT,  fill=tk.Y)
hsb.pack(side=tk.BOTTOM, fill=tk.X)
tree_disp.pack(fill="both", expand=True)

# ── Alternating row colours ──────────────────────────────────────────────────
tree_disp.tag_configure("odd",  background="#F5F5F5")
tree_disp.tag_configure("even", background="#FFFFFF")

# patch display_max_displacements to apply row tags
_orig_display = display_max_displacements
def display_max_displacements(disp_rows):
    for row in tree_disp.get_children():
        tree_disp.delete(row)
    for idx, r in enumerate(disp_rows):
        tag = "even" if idx % 2 == 0 else "odd"
        tree_disp.insert("", tk.END, values=(
            r["subcase"],
            f"{r['T1']:.2f}",
            f"{r['T2']:.2f}",
            f"{r['T3']:.2f}",
            f"{r['R1']:.2f}",
            f"{r['R2']:.2f}",
            f"{r['R3']:.2f}",
        ), tags=(tag,))
    lbl_disp_count.config(text=f"{len(disp_rows)} row(s) found")

root.mainloop()
