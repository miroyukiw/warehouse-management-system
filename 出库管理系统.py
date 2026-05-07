import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import datetime
import re

def connect_db():
    return sqlite3.connect("warehouse.db")

def get_distinct_from_outtable(field):
    """
    从“出库表”里获取某字段(如'出库日期','领用人','经办人','项目名称','项目编码')的所有distinct值
    """
    conn = connect_db()
    c = conn.cursor()
    data = []
    try:
        c.execute(f"SELECT DISTINCT {field} FROM 出库表 WHERE {field} IS NOT NULL")
        data = [r[0] for r in c.fetchall() if r[0]]
    except:
        pass
    conn.close()
    return data

def show_inventory_window():
    win = tk.Toplevel(root)
    win.title("出库管理系统 - 多功能整合")
    win.geometry("1000x600")
    win.lift()        # 保持在最前
    win.focus_force() # 防止最小化

    style = ttk.Style(win)
    style.theme_use("clam")
    style.configure("Line.Treeview",
                    background="white",
                    rowheight=30,
                    bordercolor="gray",
                    borderwidth=1,
                    relief="solid")
    style.configure("Line.Treeview.Heading",
                    background="#f0f0f0",
                    bordercolor="gray")

    # 顶部搜索区
    top_frame = ttk.Frame(win)
    top_frame.pack(side="top", fill="x")

    ttk.Label(top_frame, text="搜索:").pack(side="left", padx=5)
    search_var = tk.StringVar()
    search_entry = ttk.Entry(top_frame, textvariable=search_var, width=20)
    search_entry.pack(side="left", padx=5)

    # “刷新”按钮 (只有一个)
    def do_refresh():
        selected_state.clear()
        checked_order.clear()
        for iid in tree.get_children():
            tree.delete(iid)
        load_data()
        do_filter("")
        win.lift()
        win.focus_force()

    btn_refresh = ttk.Button(top_frame, text="刷新", command=do_refresh)
    btn_refresh.pack(side="right", padx=10)

    # 中间 Treeview
    main_frame = ttk.Frame(win)
    main_frame.pack(side="top", expand=True, fill="both")

    v_scroll = ttk.Scrollbar(main_frame, orient="vertical")
    h_scroll = ttk.Scrollbar(main_frame, orient="horizontal")

    columns = ("chk", "物资名称", "规格型号", "厂家", "单位", "数量", "plus")
    tree = ttk.Treeview(main_frame,
                        columns=columns,
                        show="headings",
                        selectmode="none",
                        yscrollcommand=v_scroll.set,
                        xscrollcommand=h_scroll.set,
                        style="Line.Treeview")
    for col, w_val, cap in zip(columns, [50, 200, 200, 180, 150, 80, 50],
                               ["勾选", "物资名称", "规格型号", "厂家", "单位", "数量", "展开"]):
        tree.heading(col, text=cap, anchor="center")
        tree.column(col, width=w_val, anchor="center")

    tree.grid(row=0, column=0, sticky="nsew")
    v_scroll.config(command=tree.yview)
    v_scroll.grid(row=0, column=1, sticky="ns")
    h_scroll.config(command=tree.xview)
    h_scroll.grid(row=1, column=0, sticky="ew")

    main_frame.rowconfigure(0, weight=1)
    main_frame.columnconfigure(0, weight=1)

    def _on_mousewheel(e):
        factor = 1.5
        if e.state & 0x0001:
            tree.xview_scroll(int(-1 * (e.delta * factor / 120)), "units")
        else:
            tree.yview_scroll(int(-1 * (e.delta * factor / 120)), "units")
        return "break"
    tree.bind("<MouseWheel>", _on_mousewheel)

    # 下方出库按钮(正中)
    bottom_frame = ttk.Frame(win)
    bottom_frame.pack(side="bottom", fill="x")

    # 数据结构定义
    all_rows = {}       # { iid -> rowvals }，存储主记录
    child_data = {}     # { (物资名称,规格型号): [(批次库存表序号, 批次或卷号, 剩余数量, 备注)] }
    selected_state = {} # { iid -> 是否被选中 }
    checked_order = []  # 记录勾选顺序
    expanded_state = {} # { iid -> 是否已展开 }
    expanded_children = {} # { iid -> [子记录id列表] }

    #####################################
    # 载入数据库，构造all_rows和child_data
    #####################################
    def load_data():
        all_rows.clear()
        child_data.clear()
        selected_state.clear()
        checked_order.clear()
        expanded_state.clear()
        expanded_children.clear()

        conn = connect_db()
        c = conn.cursor()
        # 先加载子表数据（批次库存表）
        brows = c.execute("""
            SELECT 序号, 物资名称, 规格型号, 批次或卷号, 剩余数量, 备注
            FROM 批次库存表
            WHERE 剩余数量 > 0
        """).fetchall()
        from collections import defaultdict
        dd = defaultdict(list)
        for (bid, nm, sp, bc, remain, rm) in brows:
            dd[(nm, sp)].append((bid, bc, remain, rm))
        child_data.update(dd)
        
        # 再加载主表数据（库存表）
        rows = c.execute("""
            SELECT 序号, 物资名称, 规格型号, 厂家, 单位, 当前库存
            FROM 库存表
            WHERE 当前库存 > 0
        """).fetchall()
        for (mid, nm, sp, fac, unt, stk) in rows:
            iid = f"main-{mid}@@fac={fac}@@unt={unt}"
            cchk = "□"
            plus = "＋" if is_has_batch(nm, sp) else ""
            rowvals = (cchk, nm, sp, fac, unt, str(stk), plus)
            all_rows[iid] = rowvals
        conn.close()

    def is_has_batch(nm, sp):
        """判断 (物资名称,规格型号) 是否在 child_data 中有记录"""
        return bool(child_data.get((nm, sp)))

    #####################################
    # 根据搜索关键字过滤并显示数据
    #####################################
    def do_filter(keyword):
        for iid in tree.get_children():
            tree.delete(iid)
        kw = keyword.lower()
        for iid, vals in all_rows.items():
            row_str = " ".join(vals)
            if kw in row_str or selected_state.get(iid, False):
                tree.insert("", "end", iid=iid, values=vals)
                expanded_state[iid] = False
                expanded_children[iid] = []

    def on_search(*args):
        do_filter(search_var.get())

    search_entry.bind("<KeyRelease>", on_search)

    #####################################
    # 切换展开/收起功能
    #####################################
    def toggle_expand(iid):
        if not tree.exists(iid):
            return
        rv = tree.item(iid, "values")
        nm = rv[1]
        sp = rv[2]
        plus = rv[6]

        # 解析 iid 中的厂家和单位信息
        parts = iid.split("@@")
        main_id = parts[0]  # 形如 "main-xx"
        fac = ""
        unt = ""
        for p in parts[1:]:
            if p.startswith("fac="):
                fac = p.replace("fac=", "")
            elif p.startswith("unt="):
                unt = p.replace("unt=", "")

        if expanded_state.get(iid, False):
            plus = "＋"
            for cc in expanded_children[iid]:
                if tree.exists(cc):
                    tree.delete(cc)
                if cc in selected_state:
                    del selected_state[cc]
            expanded_children[iid] = []
            expanded_state[iid] = False
        else:
            plus = "－"
            expanded_state[iid] = True
            if (nm, sp) in child_data:
                pindex = tree.index(iid)
                kids = []
                for (bid, bc, remain, rm) in child_data[(nm, sp)]:
                    child_iid = f"batch-{bid}@@fac={fac}@@unt={unt}@@parentid={main_id}"
                    cchk = "□"
                    cvals = (cchk, nm, sp, bc, rm, str(remain), "")
                    tree.insert("", pindex + 1, iid=child_iid, values=cvals)
                    pindex += 1
                    kids.append(child_iid)
                expanded_children[iid] = kids

        newvals = (rv[0], rv[1], rv[2], rv[3], rv[4], rv[5], plus)
        tree.item(iid, values=newvals)

    #####################################
    # 切换选择状态（勾选）
    #####################################
    def toggle_check(iid):
        if not tree.exists(iid):
            return
        old = selected_state.get(iid, False)
        newv = not old
        selected_state[iid] = newv
        rv = tree.item(iid, "values")
        cchk = "☑" if newv else "□"
        newvals = (cchk,) + rv[1:]
        tree.item(iid, values=newvals)
        if newv:
            if iid not in checked_order:
                checked_order.append(iid)
        else:
            if iid in checked_order:
                checked_order.remove(iid)

    #####################################
    # 自定义输入对话框，询问出库数量
    #####################################
    def ask_quantity_dialog(parent, title, prompt, default_val=""):
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.lift()
        dialog.attributes("-topmost", True)
        dialog.focus_force()

        frm = ttk.Frame(dialog, padding=10)
        frm.pack(expand=True, fill="both")

        ttk.Label(frm, text=prompt).pack(side="top", padx=5, pady=5)

        var = tk.StringVar(value=default_val)
        entry = ttk.Entry(frm, textvariable=var, width=15)
        entry.pack(side="top", padx=5, pady=5)
        entry.focus()

        result = [None]

        def on_ok(e=None):
            result[0] = var.get().strip()
            dialog.destroy()

        def on_cancel(e=None):
            result[0] = None
            dialog.destroy()

        btn_frame = ttk.Frame(frm)
        btn_frame.pack(side="top", pady=5)

        ok_btn = ttk.Button(btn_frame, text="OK", command=on_ok)
        ok_btn.pack(side="left", padx=5)
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=on_cancel)
        cancel_btn.pack(side="left", padx=5)

        entry.bind("<Return>", on_ok)
        dialog.bind("<Escape>", on_cancel)

        dialog.grab_set()
        parent.wait_window(dialog)
        return result[0]

    #####################################
    # 出库操作：遍历勾选记录进行处理
    #####################################
    def do_out():
        if not any(selected_state.values()):
            messagebox.showinfo("提示", "请至少勾选一行!")
            return

        out_dlg = tk.Toplevel(win)
        out_dlg.title("出库信息")
        out_dlg.lift()
        out_dlg.attributes("-topmost", True)
        out_dlg.focus_force()

        frm = ttk.Frame(out_dlg, padding=10)
        frm.pack(expand=True, fill='both')

        out_fields = ["出库日期", "领用人", "经办人", "项目名称", "项目编码"]
        combos = {}
        data_out = {}
        for f in out_fields:
            data_out[f] = get_distinct_from_outtable(f)

        for i, f in enumerate(out_fields):
            lb = ttk.Label(frm, text=f, width=12, anchor='e')
            lb.grid(row=i, column=0, padx=5, pady=5, sticky='e')
            cb = ttk.Combobox(frm, width=25, values=data_out[f])
            cb.grid(row=i, column=1, padx=5, pady=5, sticky='w')
            combos[f] = cb

        def confirm_out():
            raw_date = combos["出库日期"].get().strip()
            if not raw_date:
                messagebox.showwarning("警告", "出库日期不可为空!")
                return
            pat = re.compile(r'^\d{8}$')
            try:
                if pat.match(raw_date):
                    dt = datetime.datetime.strptime(raw_date, "%Y%m%d")
                    out_date = dt.strftime("%Y/%m/%d")
                else:
                    out_date = raw_date
            except:
                messagebox.showwarning("警告", "日期格式错误!")
                return

            rec = combos["领用人"].get().strip()
            opr = combos["经办人"].get().strip()
            pnm = combos["项目名称"].get().strip()
            pcd = combos["项目编码"].get().strip()

            conn = connect_db()
            c = conn.cursor()

            current_list = checked_order[:]
            for iid in current_list:
                if not selected_state.get(iid, False):
                    continue
                if not tree.exists(iid):
                    continue
                vals = tree.item(iid, "values")
                nm = vals[1]
                sp = vals[2]

                parts = iid.split("@@")
                item_part = parts[0]  # 例如 "main-12" 或 "batch-99"
                fac = ""
                unt = ""
                parentid = ""
                for p in parts[1:]:
                    if p.startswith("fac="):
                        fac = p.replace("fac=", "")
                    elif p.startswith("unt="):
                        unt = p.replace("unt=", "")
                    elif p.startswith("parentid="):
                        parentid = p.replace("parentid=", "")

                if iid.startswith("main-"):
                    stock_str = vals[5]
                    try:
                        stock_int = int(stock_str)
                    except:
                        stock_int = 0
                    main_id = item_part.replace("main-", "")
                    try:
                        mid = int(main_id)
                    except:
                        mid = 0
                    prompt = f"【主行】\n物资:{nm}\n规格:{sp}\n厂家:{fac}\n单位:{unt}\n库存:{stock_int}\n出库数量?"
                    qty_str = ask_quantity_dialog(out_dlg, "出库数量", prompt)
                    if qty_str is None:
                        continue
                    try:
                        out_qty = int(qty_str)
                    except:
                        messagebox.showwarning("警告", "数字无效!")
                        continue
                    if out_qty <= 0 or out_qty > stock_int:
                        messagebox.showwarning("警告", f"数量不在合法范围(1~{stock_int})")
                        continue
                    new_st = stock_int - out_qty
                    c.execute("""
                        UPDATE 库存表
                        SET 当前库存 = ?, 最近出库日期 = ?
                        WHERE 序号 = ?
                    """, (new_st, out_date, mid))
                    c.execute("""
                        INSERT INTO 出库表(
                            出库日期,物资名称,规格型号,厂家,
                            单位,数量,领用人,经办人,
                            项目名称,项目编码
                        )
                        VALUES(?,?,?,?,?,?,?,?,?,?)
                    """, (
                        out_date, nm, sp, fac, unt, out_qty,
                        rec, opr, pnm, pcd
                    ))
                    if new_st <= 0:
                        tree.delete(iid)
                        if iid in selected_state:
                            selected_state.pop(iid)
                        if iid in checked_order:
                            checked_order.remove(iid)
                    else:
                        newvals = (vals[0], nm, sp, fac, unt, str(new_st), vals[6])
                        tree.item(iid, values=newvals)

                elif iid.startswith("batch-"):
                    remain_str = vals[5]
                    try:
                        remain_int = int(remain_str)
                    except:
                        remain_int = 0
                    out_qty = remain_int
                    batch_id = item_part.replace("batch-", "")
                    try:
                        bid = int(batch_id)
                    except:
                        bid = 0
                    c.execute("""
                        UPDATE 批次库存表
                        SET 剩余数量 = 0, 最近出库日期 = ?
                        WHERE 序号 = ?
                    """, (out_date, bid))

                    if parentid.startswith("main-"):
                        pid = parentid.replace("main-", "")
                        try:
                            pid_int = int(pid)
                        except:
                            pid_int = 0
                        prow = c.execute("""
                            SELECT 当前库存 FROM 库存表
                            WHERE 序号 = ?
                        """, (pid_int,)).fetchone()
                        if prow:
                            pstock = prow[0]
                        else:
                            pstock = 0
                        newp = pstock - out_qty
                        if newp < 0:
                            newp = 0
                        c.execute("""
                            UPDATE 库存表
                            SET 当前库存 = ?, 最近出库日期 = ?
                            WHERE 序号 = ?
                        """, (newp, out_date, pid_int))

                        for topiid in tree.get_children():
                            if topiid.startswith(parentid):
                                if not tree.exists(topiid):
                                    break
                                pvals = tree.item(topiid, "values")
                                try:
                                    oldp = int(pvals[5])
                                except:
                                    oldp = 0
                                np = oldp - out_qty
                                if np < 0:
                                    np = 0
                                if np == 0:
                                    tree.delete(topiid)
                                    if topiid in selected_state:
                                        selected_state.pop(topiid)
                                    if topiid in checked_order:
                                        checked_order.remove(topiid)
                                else:
                                    newvals = (pvals[0], pvals[1], pvals[2],
                                               pvals[3], pvals[4], str(np), pvals[6])
                                    tree.item(topiid, values=newvals)
                                break

                    c.execute("""
                        INSERT INTO 出库表(
                            出库日期,物资名称,规格型号,厂家,
                            单位,数量,领用人,经办人,
                            项目名称,项目编码
                        )
                        VALUES(?,?,?,?,?,?,?,?,?,?)
                    """, (out_date, nm, sp, fac, unt, out_qty, rec, opr, pnm, pcd))
                    tree.delete(iid)
                    if iid in selected_state:
                        selected_state.pop(iid)
                    if iid in checked_order:
                        checked_order.remove(iid)

            conn.commit()
            conn.close()

            messagebox.showinfo("成功", "出库完成！")
            out_dlg.destroy()
            # ★新加：出库完成后刷新界面，以便清除勾选并重载数据
            do_refresh()

        btn_confirm = ttk.Button(frm, text="确认出库", command=confirm_out)
        btn_confirm.grid(row=len(out_fields), column=0, columnspan=2, pady=10)

    btn_out = ttk.Button(bottom_frame, text="出库", command=do_out)
    btn_out.pack(side="top", pady=10)

    # 点击事件：根据点击位置判断是展开/收起还是切换勾选状态
    def on_click(e):
        region = tree.identify("region", e.x, e.y)
        if region != "cell":
            return
        col_id = tree.identify_column(e.x)
        row_id = tree.identify_row(e.y)
        if not row_id:
            return
        cindex = int(col_id.replace("#", "")) - 1
        if not tree.exists(row_id):
            return
        vals = tree.item(row_id, "values")
        if cindex == 6:
            if vals[6] in ["＋", "－"]:
                toggle_expand(row_id)
        else:
            toggle_check(row_id)

    tree.bind("<Button-1>", on_click)

    def load_all():
        load_data()
        do_filter("")

    load_all()

def main():
    global root
    root.title("主窗口")
    root.geometry("400x300")
    ttk.Button(root, text="打开出库管理", command=show_inventory_window).pack(expand=True, pady=30)
    root.mainloop()

root = tk.Tk()
main()
