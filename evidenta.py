import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3, csv, os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import date, datetime

APP_DIR=os.path.join(os.environ.get('APPDATA',os.path.expanduser('~')),'EvidentaDocumente')
os.makedirs(APP_DIR,exist_ok=True)
DB=os.path.join(APP_DIR,'evidenta.db')
SERVICES=['Inspecție tahograf','Atestat profesional','ADR','CPI','CPC','Altele']
RED='#b52b2b'; BLUE='#294c91'; WHITE='#ffffff'; BG='#f2f4f7'

def connect():
    c=sqlite3.connect(DB)
    c.execute('''CREATE TABLE IF NOT EXISTS records (id INTEGER PRIMARY KEY AUTOINCREMENT, data_inregistrarii TEXT DEFAULT '', nume TEXT, prenume TEXT, telefon TEXT, serviciu TEXT, auto TEXT, efectuare TEXT, expirare TEXT)''')
    cols=[x[1] for x in c.execute('PRAGMA table_info(records)').fetchall()]
    if 'data_inregistrarii' not in cols:
        c.execute("ALTER TABLE records ADD COLUMN data_inregistrarii TEXT DEFAULT ''")
        c.execute("UPDATE records SET data_inregistrarii=? WHERE data_inregistrarii IS NULL OR data_inregistrarii=''",(date.today().strftime('%d.%m.%Y'),))
    if 'nr_referinta' not in cols:
        c.execute("ALTER TABLE records ADD COLUMN nr_referinta INTEGER")
        c.execute("UPDATE records SET nr_referinta=id WHERE nr_referinta IS NULL")
    if 'observatii' not in cols:
        c.execute("ALTER TABLE records ADD COLUMN observatii TEXT DEFAULT ''")
    c.execute("""
        CREATE TABLE IF NOT EXISTS istoric_modificari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_id INTEGER NOT NULL,
            date_vechi TEXT NOT NULL,
            data_modificarii TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    c.commit();return c

class App:
 def __init__(self,root):
    self.root=root;root.title('TTC | Evidență Atestate și Inspecții');root.geometry('1400x800');root.minsize(1050,650);root.configure(bg=BG)
    self.conn=connect();self.selected=None;self.build();self.refresh();self.alerts()
    self.root.after(300,lambda:self.nume_entry.focus_force())
 def build(self):
    style=ttk.Style();style.theme_use('clam');style.configure('Treeview',rowheight=28,background='white',fieldbackground='white');style.configure('Treeview.Heading',font=('Segoe UI',9,'bold'),foreground=BLUE);style.configure('TButton',font=('Segoe UI',9));style.configure('TLabel',background=BG)
    header=tk.Frame(self.root,bg=WHITE);header.pack(fill='x',padx=10,pady=(10,4))
    logo_frame=tk.Frame(header,bg=WHITE)
    logo_frame.pack(side='left',padx=(0,18))
    try:
        from PIL import Image,ImageTk
        im=Image.open(os.path.join(getattr(sys,'_MEIPASS',os.path.dirname(__file__)),'logo.jpeg'))
        from PIL import ImageChops
        bbox=ImageChops.difference(im.convert('RGB'),Image.new('RGB',im.size,'white')).getbbox()
        if bbox:
            im=im.crop(bbox)
        im.thumbnail((250,85))
        self.logo=ImageTk.PhotoImage(im)
        tk.Label(logo_frame,image=self.logo,bg=WHITE).pack()
        tk.Label(logo_frame,text='COMPANY',font=('Segoe UI',12,'bold'),fg=BLUE,bg=WHITE).pack(pady=(0,0),anchor='center')
    except Exception as e:
        messagebox.showerror('Eroare siglă', str(e))
        tk.Label(header,text='TTC',font=('Segoe UI',26,'bold'),fg=RED,bg=WHITE).pack(side='left',padx=18)
    tk.Label(header,text='EVIDENȚĂ ATESTATE ȘI INSPECȚII TAHOGRAF',font=('Segoe UI',18,'bold'),fg=BLUE,bg=WHITE).pack(side='left',padx=12)
    stats=tk.Frame(self.root,bg=BG)
    stats.pack(fill='x',padx=10)

    self.statslabels={}
    self.filter_status=tk.StringVar(value='Toate')

    categories=[
        ('exp','EXPIRATE'),
        ('30','≤ 30 ZILE'),
        ('60','31–60 ZILE'),
        ('90','61–90 ZILE'),
        ('incomplete','NECOMPLET'),
        ('all','TOATE')
    ]

    for key,label in categories:
        box=ttk.LabelFrame(stats,text=label)
        box.pack(side='left',fill='x',expand=True,padx=5)

        lbl=ttk.Label(box,text='0',font=('Segoe UI',18,'bold'),foreground=RED)
        lbl.pack(pady=6)

        self.statslabels[key]=lbl
        lbl.bind('<Button-1>',lambda e,k=key:self.set_status_filter(k))
        box.bind('<Button-1>',lambda e,k=key:self.set_status_filter(k))

    form=ttk.LabelFrame(self.root,text='Date persoană / document');form.pack(fill='x',padx=10,pady=8)
    keys=['ref','data_inregistrarii','nume','prenume','telefon','serviciu','auto','efectuare','expirare','observatii']
    self.vars={k:tk.StringVar() for k in keys};self.vars['data_inregistrarii'].set(date.today().strftime('%d.%m.%Y'))
    fields=[('Nr. înregistrare','ref'),('Data înregistrării','data_inregistrarii'),('Nume','nume'),('Prenume','prenume'),('Telefon','telefon'),('Serviciu','serviciu'),('Număr auto','auto'),('Data efectuării (ZZ.LL.AAAA)','efectuare'),('Data expirării (ZZ.LL.AAAA)','expirare'),('Observații','observatii')]
    for i,(label,key) in enumerate(fields):
      row=i//5;col=(i%5)*2;ttk.Label(form,text=label).grid(row=row,column=col,padx=4,pady=4,sticky='w')
      if key=='serviciu':w=ttk.Combobox(form,textvariable=self.vars[key],values=SERVICES,state='normal',width=19)
      else:w=ttk.Entry(form,textvariable=self.vars[key],width=19)
      if key=='nume':
        self.nume_entry=w
      w.grid(row=row,column=col+1,padx=4,pady=4,sticky='ew')
    buttons=ttk.Frame(form);buttons.grid(row=2,column=0,columnspan=10,pady=5)
    for label,cmd in [('Adaugă / Salvează',self.save),('Șterge',self.delete),('Curăță formular',self.clear),('Export Excel (CSV)',self.export)]:ttk.Button(buttons,text=label,command=cmd).pack(side='left',padx=5)
    searchrow=ttk.Frame(self.root);searchrow.pack(fill='x',padx=10,pady=5)
    ttk.Label(searchrow,text='Caută:').pack(side='left');self.search=tk.StringVar();ttk.Entry(searchrow,textvariable=self.search,width=35).pack(side='left',padx=5);self.search.trace_add('write',lambda *_:self.refresh())
    ttk.Label(searchrow,text='Filtru serviciu:').pack(side='left',padx=(15,0));self.filter=tk.StringVar(value='Toate');cb=ttk.Combobox(searchrow,textvariable=self.filter,values=['Toate']+SERVICES,state='readonly',width=22);cb.pack(side='left',padx=5);cb.bind('<<ComboboxSelected>>',lambda e:self.refresh())
    cols=('ref','data_inregistrarii','nume','prenume','telefon','serviciu','auto','efectuare','expirare','status','observatii');self.tree=ttk.Treeview(self.root,columns=cols,show='headings',selectmode='browse')
    heads={'ref':'Nr. înregistrare','data_inregistrarii':'Data înregistrării','nume':'Nume','prenume':'Prenume','telefon':'Telefon','serviciu':'Serviciu','auto':'Nr. auto','efectuare':'Data efectuării','expirare':'Data expirării','status':'Status','observatii':'Observații'}
    widths={'ref':150,'data_inregistrarii':120,'nume':110,'prenume':110,'telefon':110,'serviciu':170,'auto':100,'efectuare':120,'expirare':120,'status':140,'observatii':250}
    for c in cols:self.tree.heading(c,text=heads[c]);self.tree.column(c,width=widths[c],minwidth=widths[c],stretch=False,anchor='center')

    scroll_x=ttk.Scrollbar(self.tree.master,orient='horizontal',command=self.tree.xview)
    self.tree.configure(xscrollcommand=scroll_x.set)
    scroll_x.pack(side='bottom',fill='x',padx=10)
    self.tree.pack(fill='both',expand=True,padx=10,pady=8);self.tree.bind('<<TreeviewSelect>>',self.select)
    self.tree.bind('<Double-1>',self.show_observatii)
    self.tree.tag_configure('expired',background='#ffc5c5');self.tree.tag_configure('soon',background='#ffe0b2');self.tree.tag_configure('later',background='#dcefdc')
    tk.Label(self.root,text='TTC | Date salvate local. Exportă periodic o copie CSV pentru backup.',font=('Segoe UI',9),bg=BG,fg=BLUE).pack(pady=5)
 def set_status_filter(self,key):
    mapping={
        'exp':'EXPIRATE',
        '30':'≤ 30 ZILE',
        '60':'31–60 ZILE',
        '90':'61–90 ZILE',
        'incomplete':'NECOMPLET',
        'all':'Toate'
    }
    self.filter_status.set(key)
    self.refresh()
 def parsed(self,s):return datetime.strptime(s.strip(),'%d.%m.%Y').date()
 def save(self):
    vals=[self.vars[k].get().strip() for k in ['data_inregistrarii','nume','prenume','telefon','serviciu','auto','efectuare','expirare','observatii']]

    vals[1]=vals[1].upper()
    vals[2]=vals[2].upper() 
     
    if not vals[1] or not vals[2] or not vals[4]:
        messagebox.showwarning('Date incomplete','Completează numele, prenumele și serviciul.')
        return

    for index, label in [(0,'Data înregistrării'),(6,'Data efectuării'),(7,'Data expirării')]:
        if vals[index]:
            try:
                self.parsed(vals[index])
            except ValueError:
                messagebox.showerror('Dată invalidă',f'{label} trebuie introdusă în format ZZ.LL.AAAA.')
                return

    if not self.selected:
        numere=self.conn.execute(
            'SELECT nr_referinta FROM records WHERE nr_referinta IS NOT NULL ORDER BY nr_referinta'
        ).fetchall()

        nr_nou=1
        for (nr,) in numere:
            if nr==nr_nou:
                nr_nou+=1
            elif nr>nr_nou:
                break
    if self.selected:
        self.conn.execute(
            'UPDATE records SET data_inregistrarii=?,nume=?,prenume=?,telefon=?,serviciu=?,auto=?,efectuare=?,expirare=?,observatii=? WHERE id=?',
            (*vals,self.selected)
        )
    else:
        self.conn.execute(
            'INSERT INTO records(nr_referinta,data_inregistrarii,nume,prenume,telefon,serviciu,auto,efectuare,expirare,observatii) VALUES(?,?,?,?,?,?,?,?,?,?)',
            (nr_nou,*vals)
        )

    self.conn.commit()
    self.clear()
    self.root.after(100, lambda: self.nume_entry.focus_force())
    self.refresh()
 def show_observatii(self,event=None):
        selectie=self.tree.selection()
        if not selectie:
            return

        valori=self.tree.item(selectie[0],'values')
        observatii=valori[-1]

        fereastra=tk.Toplevel(self.root)
        fereastra.title('TTC | Observații')
        fereastra.geometry('650x400')
        fereastra.transient(self.root)

        ttk.Label(
            fereastra,
            text='OBSERVAȚII COMPLETE',
            font=('Segoe UI',14,'bold')
        ).pack(pady=10)

        zona=ttk.Frame(fereastra)
        zona.pack(fill='both',expand=True,padx=15,pady=10)

        text=tk.Text(zona,wrap='word',font=('Segoe UI',11))
        bara=ttk.Scrollbar(zona,orient='vertical',command=text.yview)
        text.configure(yscrollcommand=bara.set)

        bara.pack(side='right',fill='y')
        text.pack(side='left',fill='both',expand=True)

        text.insert('1.0',observatii)
        text.configure(state='disabled')

        ttk.Button(
            fereastra,
            text='Închide',
            command=fereastra.destroy
        ).pack(pady=10)
 def clear(self):
    self.selected=None
    for v in self.vars.values():v.set('')
    self.vars['data_inregistrarii'].set(date.today().strftime('%d.%m.%Y'));self.tree.selection_remove(self.tree.selection())
 def select(self,event=None):
    s=self.tree.selection()
    if not s:return
    self.selected=int(s[0]);row=self.conn.execute('SELECT id,data_inregistrarii,nume,prenume,telefon,serviciu,auto,efectuare,expirare,observatii FROM records WHERE id=?',(self.selected,)).fetchone()
    self.vars['ref'].set(str(self.conn.execute('SELECT nr_referinta FROM records WHERE id=?',(self.selected,)).fetchone()[0]))
    for k,v in zip(['data_inregistrarii','nume','prenume','telefon','serviciu','auto','efectuare','expirare','observatii'],row[1:]):self.vars[k].set(v or '')
 def delete(self):
    if not self.selected:messagebox.showinfo('Selectare','Selectează un rând din tabel.');return
    if messagebox.askyesno('Confirmare','Ștergi înregistrarea selectată?'):self.conn.execute('DELETE FROM records WHERE id=?',(self.selected,));self.conn.commit();self.clear();self.refresh()
 def refresh(self):
    for i in self.tree.get_children():
        self.tree.delete(i)

    today=date.today()
    counts={'exp':0,'30':0,'60':0,'90':0,'incomplete':0,'all':0}
    q=self.search.get().lower()
    fil=self.filter.get()
    selected_status=self.filter_status.get()

    rows=self.conn.execute('SELECT id,data_inregistrarii,nume,prenume,telefon,serviciu,auto,efectuare,expirare,nr_referinta,observatii FROM records ORDER BY nr_referinta').fetchall()

    for r in rows:
      if r[8]:
        try:
          d=self.parsed(r[8])
          delta=(d-today).days
        except Exception:
          delta=None
      else:
        delta=None

      if delta is None:
        status='NECOMPLET'
        tag=''
        counts['incomplete']+=1
      elif delta<0:
        status='EXPIRAT'
        tag='expired'
        counts['exp']+=1
      elif delta<=30:
        status=f'Expiră în {delta} zile'
        tag='soon'
        counts['30']+=1
      elif delta<=60:
        status=f'Expiră în {delta} zile'
        tag='later'
        counts['60']+=1
      elif delta<=90:
        status=f'Expiră în {delta} zile'
        tag='later'
        counts['90']+=1
      else:
        status=f'{delta} zile'
        tag=''

      counts['all']+=1

      searchable=' '.join(str(x or '') for x in r).lower()

      if q and q not in searchable:
        continue
      if fil!='Toate' and r[5]!=fil:
        continue

      if selected_status=='exp' and not (delta is not None and delta<0):
        continue
      if selected_status=='30' and not (delta is not None and 0<=delta<=30):
        continue
      if selected_status=='60' and not (delta is not None and 31<=delta<=60):
        continue
      if selected_status=='90' and not (delta is not None and 61<=delta<=90):
        continue
      if selected_status=='incomplete' and delta is not None:
        continue

      self.tree.insert('', 'end',iid=str(r[0]),values=(r[9],*r[1:9],status,r[10] or ''),tags=(tag,) if tag else ())

    for k,v in counts.items():
      if k in self.statslabels:
        self.statslabels[k].config(text=str(v))
 def alerts(self):
    upcoming=[]

    for r in self.conn.execute('SELECT nume,prenume,serviciu,auto,expirare FROM records').fetchall():
      if not r[4]:
        continue

      try:
        d=self.parsed(r[4])
        delta=(d-date.today()).days
      except Exception:
        continue

      if 0<=delta<=30:
        upcoming.append(f'{r[0]} {r[1]} — {r[2]} ({r[3] or "fără număr auto"}), expiră la {r[4]}: {delta} zile')

    if upcoming:
      messagebox.showwarning('TTC – Atenționare expirare document','\n\n'.join(upcoming))
 def export(self):
    path=filedialog.asksaveasfilename(
        defaultextension='.xlsx',
        filetypes=[('Excel Workbook','*.xlsx')],
        title='Exportă evidența TTC în Excel'
    )

    if not path:
        return

    try:
        rows=self.conn.execute(
            '''SELECT nr_referinta,data_inregistrarii,nume,prenume,
            telefon,serviciu,auto,efectuare,expirare,observatii
            FROM records ORDER BY nr_referinta'''
        ).fetchall()

        wb=Workbook()
        ws=wb.active
        ws.title='Evidenta TTC'

        headers=[
            'Nr. înregistrare','Data înregistrării','Nume',
            'Prenume','Telefon','Serviciu','Număr auto',
            'Data efectuării','Data expirării','Observații'
        ]

        ws.append(headers)

        for row in rows:
            ws.append(list(row))

        # Antet albastru TTC
        for cell in ws[1]:
            cell.fill=PatternFill(
                start_color='294C91',
                end_color='294C91',
                fill_type='solid'
            )
            cell.font=Font(color='FFFFFF',bold=True)
            cell.alignment=Alignment(horizontal='center')

        # Dimensiunea coloanelor
        widths=[16,20,20,20,18,25,18,20,20,40]

        for i,width in enumerate(widths,1):
            ws.column_dimensions[get_column_letter(i)].width=width

        # Telefonul este text, pentru a păstra zero inițial
        for cell in ws['E'][1:]:
            cell.number_format='@'

        # Evidențierea documentelor expirate sau aproape de expirare
        today=date.today()

        for row in range(2,ws.max_row+1):
            value=ws.cell(row=row,column=9).value

            if value:
                try:
                    expiry=self.parsed(str(value))
                    days=(expiry-today).days

                    if days<0:
                        color='FFC7CE'
                    elif days<=30:
                        color='FFE0B2'
                    else:
                        color=None

                    if color:
                        for cell in ws[row]:
                            cell.fill=PatternFill(
                                start_color=color,
                                end_color=color,
                                fill_type='solid'
                            )
                except (ValueError,TypeError):
                    pass

        # Filtre și antet fix
        ws.auto_filter.ref=ws.dimensions
        ws.freeze_panes='A2'

        wb.save(path)

        messagebox.showinfo(
            'Export reușit',
            'Fișierul Excel a fost salvat cu succes.'
        )

    except Exception as e:
        messagebox.showerror('Eroare export',str(e))
if __name__=='__main__':
 root=tk.Tk();app=App(root);root.mainloop()
