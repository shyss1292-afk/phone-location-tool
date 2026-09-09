import tkinter as tk
from tkinter import ttk,filedialog,messagebox
import re, json, os
try:
 import pandas as pd
except ImportError: pd=None
import sys
BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
DB_FILE = os.path.join(BASE_DIR, 'prefix_db.json')
DEFAULT_DB={'updated':'2024-01','coverage':'中国大陆手机号前7位示例号段；请按需更新prefix_db.json','prefixes':{'1380013':['中国移动','北京','北京'],'1390013':['中国移动','北京','北京'],'1860013':['中国联通','北京','北京'],'1330013':['中国电信','北京','北京'],'1810013':['中国电信','北京','北京']}}
def load_db():
 try:return json.load(open(DB_FILE,encoding='utf8'))
 except:return DEFAULT_DB
def identify(s,db):
 s=re.sub(r'\s+','',str(s)); valid=bool(re.fullmatch(r'1[3-9]\d{9}',s))
 if not valid:return s,'','','格式错误'
 p=db['prefixes'].get(s[:7]); return (s,*p,'已识别') if p else (s,'','','号段未覆盖')
class App:
 def __init__(self,root):
  self.root=root;root.title('手机号批量归属地识别工具');root.geometry('900x560');self.rows=[];self.db=load_db()
  top=ttk.Frame(root);top.pack(fill='x',padx=10,pady=8)
  ttk.Button(top,text='导入 Excel',command=self.load).pack(side='left');ttk.Button(top,text='识别',command=self.run).pack(side='left',padx=6);ttk.Button(top,text='导出 Excel',command=self.export).pack(side='left');self.info=ttk.Label(top,text=f"数据库：{self.db.get('updated')} | {self.db.get('coverage')}");self.info.pack(side='left',padx=12)
  self.tree=ttk.Treeview(root,columns=('phone','carrier','prov','city','status'),show='headings');
  for c,t,w in [('phone','手机号',150),('carrier','运营商',130),('prov','省份',130),('city','城市',130),('status','识别状态',180)]:self.tree.heading(c,text=t);self.tree.column(c,width=w)
  self.tree.pack(fill='both',expand=True,padx=10);self.progress=ttk.Progressbar(root,mode='determinate');self.progress.pack(fill='x',padx=10,pady=6);self.tip=ttk.Label(root,text='请导入包含手机号列的 xlsx 文件');self.tip.pack(anchor='w',padx=10)
 def load(self):
  if pd is None:return messagebox.showerror('缺少依赖','请安装 pandas/openpyxl')
  f=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')]);
  if not f:return
  df=pd.read_excel(f); col=next((c for c in df.columns if any(k in str(c).lower() for k in ['手机','电话','mobile','phone'])),None)
  if col is None:
   for c in df.columns:
    if df[c].astype(str).str.match(r'1[3-9]\d{9}').any():col=c;break
  if col is None:return messagebox.showwarning('提示','未找到手机号列')
  vals=df[col].dropna().astype(str).tolist(); seen=set();self.rows=[]
  for v in vals:
   if v in seen:continue
   seen.add(v);self.rows.append((v,'','','','待识别'))
  self.refresh();self.tip.config(text=f'已导入 {len(self.rows)} 条（重复数据已去重，空值已忽略）')
 def run(self):
  self.progress['maximum']=len(self.rows);self.progress['value']=0
  out=[]
  for i,r in enumerate(self.rows):out.append(identify(r[0],self.db));self.progress['value']=i+1;self.root.update_idletasks()
  self.rows=out;self.refresh();self.tip.config(text=f'识别完成，共 {len(out)} 条')
 def refresh(self):
  self.tree.delete(*self.tree.get_children());[self.tree.insert('', 'end',values=r) for r in self.rows]
 def export(self):
  if pd is None or not self.rows:return messagebox.showwarning('提示','没有可导出的数据')
  f=filedialog.asksaveasfilename(defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')]);
  if f:pd.DataFrame(self.rows,columns=['手机号','运营商','省份','城市','识别状态']).to_excel(f,index=False);messagebox.showinfo('完成','导出成功')
root=tk.Tk();App(root);root.mainloop()
