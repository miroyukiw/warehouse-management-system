import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
import datetime, re

DB = 'warehouse.db'
SEARCH_DELAY = 300   # ms

# -------------------------------------------------
# 通用函数
# -------------------------------------------------
def sanitize(s: str) -> str:
    """去掉所有空白字符（含全角空格）"""
    return re.sub(r'\s+', '', str(s))

def connect_db():
    return sqlite3.connect(DB)

# -------------------------------------------------
# 下拉候选 & 物资→规格型号字典
# -------------------------------------------------
def get_distinct(field):
    sql = f"""
      SELECT DISTINCT REPLACE(REPLACE(TRIM({field}),'　',''),' ','')
        FROM 入库表
       WHERE {field} IS NOT NULL
         AND REPLACE(REPLACE(TRIM({field}),'　',''),' ','') <> ''
    """
    with connect_db() as con:
        return [r[0] for r in con.execute(sql)]

def load_all_dict(fields):
    return {f: get_distinct(f) for f in fields}

def build_name_specs_dict():
    d={}
    with connect_db() as con:
        sql="""
          SELECT REPLACE(REPLACE(TRIM(物资名称),'　',''),' ','') AS nm,
                 REPLACE(REPLACE(TRIM(规格型号),'　',''),' ','') AS sp
            FROM 入库表 WHERE nm<>'' AND sp<>''
          UNION
          SELECT REPLACE(REPLACE(TRIM(物资名称),'　',''),' ',''),
                 REPLACE(REPLACE(TRIM(规格型号),'　',''),' ','')
            FROM 库存表
           WHERE REPLACE(REPLACE(TRIM(物资名称),'　',''),' ','')<>'' AND
                 REPLACE(REPLACE(TRIM(规格型号),'　',''),' ','')<>''"""
        for nm,sp in con.execute(sql):
            d.setdefault(nm,set()).add(sp)
    return {k:sorted(v) for k,v in d.items()}

# -------------------------------------------------
# 实时搜索
# -------------------------------------------------
search_jobs={}          # Combobox -> after_id
combobox_to_field={}
def delayed_search(cb: ttk.Combobox, char:str):
    if char=='\x08':         # Backspace 不触发
        return
    if cb in search_jobs:
        cb.after_cancel(search_jobs[cb])
    search_jobs[cb]=cb.after(SEARCH_DELAY, lambda w=cb:_do_search(w))
def _do_search(cb):
    field=combobox_to_field.get(cb,'')
    full=all_data_dict.get(field,[])
    typed=sanitize(cb.get()).lower()
    cb['values']=[v for v in full if typed in v.lower()]

def cancel_pending_search(cb):
    if cb in search_jobs:
        cb.after_cancel(search_jobs[cb])
        del search_jobs[cb]

# -------------------------------------------------
# 写库（保持不变，代码略长，这里完整保留）
# -------------------------------------------------
def insert_record(rec: dict, write_remark: bool):
    with connect_db() as con:
        cur=con.cursor()
        # 1 入库表
        cur.execute("""
           INSERT INTO 入库表(
             入库日期,物资名称,规格型号,专业类别,单位,数量,批次或卷号,厂家,
             领料人,验收人,委托人,备注,项目名称,项目编码,
             物资白单单号,物资属性,明细账存货编码,物资黄单单号)
           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,(
            rec['入库日期'],rec['物资名称'],rec['规格型号'],rec['专业类别'],
            rec['单位'],rec['数量'],rec['批次或卷号'],rec['厂家'],
            rec['领料人'],rec['验收人'],rec['委托人'],rec['备注'],
            rec['项目名称'],rec['项目编码'],
            rec['物资白单单号'],rec['物资属性'],
            rec['明细账存货编码'],rec['物资黄单单号']))
        # 2 库存表
        rk=rec['备注'] if write_remark else None
        cur.execute("""
          SELECT 序号 FROM 库存表
           WHERE REPLACE(REPLACE(TRIM(物资名称),'　',''),' ','')=?
             AND REPLACE(REPLACE(TRIM(规格型号),'　',''),' ','')=?
             AND TRIM(IFNULL(厂家,''))=?
             AND (备注=? OR (备注 IS NULL AND ? IS NULL))
        """,(sanitize(rec['物资名称']),sanitize(rec['规格型号']),
             rec['厂家'],rk,rk))
        row=cur.fetchone()
        if row:
            cur.execute("""
              UPDATE 库存表
                 SET 当前库存 = 当前库存 + ?,
                     最近入库日期=?, 厂家=?, 备注=?
               WHERE 序号=?""",
               (rec['数量'],rec['入库日期'],rec['厂家'],rk,row[0]))
        else:
            cur.execute("""
              INSERT INTO 库存表(
                物资名称,规格型号,专业类别,委托人,厂家,单位,
                当前库存,最近入库日期,备注)
              VALUES(?,?,?,?,?,?,?,?,?)""",
              (rec['物资名称'],rec['规格型号'],rec['专业类别'],
               rec['委托人'],rec['厂家'],rec['单位'],
               rec['数量'],rec['入库日期'],rk))
        # 3 批次库存表
        if rec['批次或卷号']:
            cur.execute("""
              SELECT 序号 FROM 批次库存表
               WHERE REPLACE(REPLACE(TRIM(物资名称),'　',''),' ','')=?
                 AND REPLACE(REPLACE(TRIM(规格型号),'　',''),' ','')=?
                 AND 批次或卷号=?""",
               (sanitize(rec['物资名称']),sanitize(rec['规格型号']),rec['批次或卷号']))
            r=cur.fetchone()
            if r:
                cur.execute("""
                  UPDATE 批次库存表
                     SET 剩余数量 = 剩余数量 + ?,
                         最近入库日期=?, 备注=?
                   WHERE 序号=?""",
                   (rec['数量'],rec['入库日期'],rec['备注'],r[0]))
            else:
                cur.execute("""
                  INSERT INTO 批次库存表(
                    物资名称,规格型号,批次或卷号,厂家,
                    剩余数量,最近入库日期,备注)
                  VALUES(?,?,?,?,?,?,?)""",
                  (rec['物资名称'],rec['规格型号'],rec['批次或卷号'],
                   rec['厂家'],rec['数量'],rec['入库日期'],rec['备注']))
        con.commit()

# -------------------------------------------------
# GUI 回调
# -------------------------------------------------
last_material=''   # 记录上一次已处理的物资名称(已 sanitize)
def refresh_all_dicts():
    global all_data_dict,name_specs
    all_data_dict = load_all_dict(combo_fields)
    name_specs = build_name_specs_dict()
    for f in combo_fields:
        cb=widgets[f]
        if isinstance(cb,ttk.Combobox):
            cb['values']=all_data_dict.get(f,[])
def on_material_change(_=None):
    global last_material
    raw = widgets['物资名称'].get()
    key = sanitize(raw)
    if key == last_material:   # 未变化 → 不做刷新
        return
    last_material = key
    specs = name_specs.get(key, [])
    cb_spec = widgets['规格型号']
    if isinstance(cb_spec, ttk.Combobox):
        cb_spec['values'] = specs
        cb_spec.set('')
    for fld in ('专业类别','单位'):
        widgets[fld].delete(0,tk.END)
    # 取消 pending 搜索，避免错位
    cancel_pending_search(widgets['物资名称'])
def on_spec_selected(_=None):
    mat=sanitize(widgets['物资名称'].get())
    spec=sanitize(widgets['规格型号'].get())
    with connect_db() as con:
        row=con.execute("""
          SELECT 单位,专业类别 FROM 入库表
           WHERE REPLACE(REPLACE(TRIM(物资名称),'　',''),' ','')=? AND
                 REPLACE(REPLACE(TRIM(规格型号),'　',''),' ','')=?
           LIMIT 1""",(mat,spec)).fetchone()
    if row:
        widgets['单位'].delete(0,tk.END); widgets['单位'].insert(0,row[0])
        widgets['专业类别'].delete(0,tk.END); widgets['专业类别'].insert(0,row[1])
def confirm_in():
    mat=sanitize(widgets['物资名称'].get())
    if not mat:
        messagebox.showwarning("提示","物资名称为空"); return
    qty=sanitize(widgets['数量'].get())
    if not qty.isdigit():
        messagebox.showwarning("错误","数量必须为整数"); return
    rec={f:sanitize(widgets[f].get()) for f in fields}
    rec['数量']=int(qty)
    try:
        insert_record(rec,remark_var.get())
    except Exception as e:
        messagebox.showerror("错误",f"写库失败: {e}"); return
    refresh_all_dicts()
    # 批次号+1
    m=re.match(r'^(.*-)(\d+)$',rec['批次或卷号'])
    if m:
        pre,num=m.groups()
        widgets['批次或卷号'].delete(0,tk.END)
        widgets['批次或卷号'].insert(0,f"{pre}{int(num)+1}")
    remark_var.set(False)
    messagebox.showinfo("成功","入库完成")
def paste_excel():
    try: clip=root.clipboard_get()
    except: messagebox.showerror("错误","无法读取剪贴板"); return
    rows=[r for r in clip.splitlines() if r.strip()]
    if not rows:
        messagebox.showwarning("提示","剪贴板为空"); return
    if len(rows)>1:
        messagebox.showinfo("提示","检测到多行，仅取第一行")
    cols=rows[0].split('\t')
    if len(cols)<8:
        messagebox.showwarning("提示","第一行列数不足8"); return
    mapping=[('物资属性',0),('明细账存货编码',1),('物资名称',2),('规格型号',3),
             ('单位',4),('数量',5),('厂家',6),('项目名称',7)]
    for fld,idx in mapping:
        widgets[fld].delete(0,tk.END)
        widgets[fld].insert(0,sanitize(cols[idx]))
    on_material_change()

# -------------------------------------------------
# GUI 构建
# -------------------------------------------------
root=tk.Tk(); root.title("入库管理系统"); root.geometry("1280x720")
frame=ttk.Frame(root,padding=20); frame.pack(expand=True,fill='both')

fields=['入库日期','物资名称','规格型号','专业类别','单位','数量','批次或卷号','厂家',
        '领料人','验收人','委托人','备注',
        '项目名称','项目编码',
        '物资白单单号','物资属性','明细账存货编码','物资黄单单号']
combo_fields=['物资名称','规格型号','单位','专业类别','厂家',
              '项目名称','项目编码','领料人','验收人','委托人']
widgets={}
all_data_dict = load_all_dict(combo_fields)
name_specs     = build_name_specs_dict()

def make_combo(f):
    cb=ttk.Combobox(frame,values=all_data_dict.get(f,[]))
    combobox_to_field[cb]=f
    cb.bind("<KeyRelease>",lambda e,w=cb:delayed_search(w,e.char))
    return cb

for i,f in enumerate(fields):
    r,c=divmod(i,2); c*=2
    ttk.Label(frame,text=f).grid(row=r,column=c,sticky='e',padx=4,pady=3)
    w=make_combo(f) if f in combo_fields else ttk.Entry(frame)
    if f=='入库日期':
        w.insert(0,datetime.date.today().strftime('%Y/%m/%d'))
    elif f in ('领料人','验收人'):
        w.insert(0,'项元诚')
    w.grid(row=r,column=c+1,sticky='ew',padx=4,pady=3)
    widgets[f]=w

# 绑定
widgets['物资名称'].bind("<<ComboboxSelected>>", on_material_change)
widgets['物资名称'].bind("<FocusOut>", on_material_change)
widgets['规格型号'].bind("<<ComboboxSelected>>", on_spec_selected)

row=(len(fields)+1)//2
ttk.Button(frame,text="确认入库",command=confirm_in).grid(row=row,column=0,columnspan=2,pady=6)
remark_var=tk.BooleanVar(value=False)
ttk.Checkbutton(frame,text="写备注到库存表",variable=remark_var
               ).grid(row=row+1,column=0,columnspan=2)
ttk.Button(frame,text="刷新下拉列表",command=refresh_all_dicts
          ).grid(row=row+2,column=0,columnspan=2,pady=3)
ttk.Button(frame,text="Excel 粘贴入库",command=paste_excel
          ).grid(row=row+3,column=0,columnspan=2,pady=3)

frame.columnconfigure((1,3),weight=1)
root.mainloop()
