#!/usr/bin/python3

# Degur 0.1.1
# Copyright (c) 2013-2015 Shershakov Dmitry. All rights reserved.
#
# You may redistribute source and code by GPL v.3
# For otherwise use - contact by author (da-sher@yandex.ru)

# Degur - Duty list
#  - tabular visualisation
#  - calculation days, hours
#  - russian calendar form and help only, any coments also, sorry


import sys
from tkinter import *
from tkinter.ttk import Treeview, Combobox

if sys.platform=='win32':
    from ctypes import windll
    FreeConsole=windll.kernel32.FreeConsole
    AllocConsole=windll.kernel32.AllocConsole
else:
    FreeConsole=lambda:None
    AllocConsole=lambda:None

import yaml

#s='*ДН--ДН--ДН-ДН--ДН--ДН--ДН--ДН--*'
def analyse(s,holydays):
    ss=s[1:-1]
    x_days=ss.count('Д')
    x_nights=ss.count('Н')
    y=x_nights*8
    x=(x_days+x_nights)*12
    for i in range(1,10):
        x+=i*ss.count(str(i))
    if s[0]=='Н': 
        x+=8
        y+=6
    if s[-2]=='Н': 
        x-=8
        y-=6
    z=0
    for h in holydays:
        if s[h]=='Д':
            z+=12
        elif s[h]=='Н':
            z+=4
        if s[h] in '123456789':
            z+=int(s[h])
        if s[h-1]=='Н':
            z+=8
    return (x,y,z) # рабочие, ночные, праздничные

def analyse2(s,d):
    ss=s[1:-1]
    x_days=ss.count('Д')
    x_nights=ss.count('Н')+ss.count('t')
    x_days2=ss.count('D')
    x_nights2=ss.count('N')
    h_nights=(x_nights+x_nights2)*8
    h_all=(x_days+x_nights)*12+x_nights2*13+x_days2*11
    
    xx=0
    for i in range(1,10):
        xx+=i*ss.count(str(i))
    h_all+=xx+ss.count('a')*10+ss.count('К')*8+ss.count('Э')*8+\
            ss.count('Т')*2+ss.count('t')*2
    if s[0] in ['Н','N','t']: 
        h_all+=8
        h_nights+=6
    if s[-2] in ['Н','N']: 
        h_all-=8
        h_nights-=6
    h_holydays=0
    for h in d['holydays']:
        if s[h]=='Д':
            h_holydays+=12
        elif s[h]=='D':
            h_holydays+=11
        elif s[h]=='Н':
            h_holydays+=4
        elif s[h]=='N':
            h_holydays+=5
        if s[h] in '123456789': # hypotetic
            h_holydays+=int(s[h])
        if s[h] in 'a': # hypotetic
            h_holydays+=10
        if s[h] in 't': # hypotetic
            h_holydays+=6
        if s[h] in ['К','Э']:
            h_holydays+=8
        if s[h-1] in ['Н','N']:
            h_holydays+=8
    offset=d['offset']
    WH=0
    zz=0
    x_otp=0
    for i in range(1,d['days']+1):
        if s[i] in ['j','J']: zz+=8
        if s[i] in ['О'] and i not in d['holydays']: x_otp+=1
            
        if i in d['holydays']: wh=0
        elif i in d['restdays'] : wh=0
        elif i in d['shortdays'] and s[i] not in ['О','Б','У','_']: wh=7
        elif i in d['workdays']  and s[i] not in ['О','Б','У','_']: wh=8
        elif (i+offset) % 7 in [0,6]: wh=0
        elif s[i] in ['О','Б','У','_']: wh=0
        else: wh=8
        WH+=wh
    return (h_all,h_nights,h_holydays,WH,zz,xx,x_otp) # рабочие, ночные, праздничные, норма

class MyHelp():
    def __init__(self,root,data,y):
        self.root=root
        w=Toplevel(root)
        w.wm_title('Help')
        w.columnconfigure(0,weight=1)
        w.rowconfigure(0,weight=1)
        #w.bind('<Key>',self.x)
        b_hide_con=Button(master=w,text='Hide Console',command=FreeConsole)
        b_show_con=Button(master=w,text='Show Console',command=AllocConsole)
        b_hide_con.grid(row=0,column=0)
        b_show_con.grid(row=0,column=1)
        t=Text(master=w)
        t.grid(row=1,column=0,columnspan=2)
        msg='''<F1> - Help
<Control-F3> - поквартальная статистика по плану
<Shift-F3> - поквартальная статистика по табелю
<Control-F4> - статистика по плану
<Shift-F4> - статистика по табелю
<F3> - поквартальная статистика (авто)
<F4> - статистика (авто)
<F5> - табель
<F7> - отпуска (авто)
<F8> - statistic_xx (авто)
<F9> - переработка (авто)
<Control-F8> - statistic_xx план
'''
        t.insert('end',msg)

        
class Gr():
    def __init__(self,root,data,SCRY=None):
        self.data=data
        self.columns=[x for x in range(1,8)]+['day']
        root.rowconfigure(1,weight=1)
        root.columnconfigure(0,weight=1)
        root.columnconfigure(1,weight=1)
        root.columnconfigure(2,weight=1)
        f=Frame(root)
        f.columnconfigure(0,weight=1)
        f.rowconfigure(1,weight=1)
        self.v=Combobox(root)
        self.v.grid(row=0,column=0)
        self.v.bind('<<ComboboxSelected>>',self.select_ver)
        f.grid(row=1,column=0,columnspan=3,sticky=N+S)
        self.tree=Treeview(f,
                columns=self.columns,
                displaycolumns=['day']+self.columns[:-1],
                show='headings')
        #self.tree.tag_configure('odd',background='white')
        #self.tree.tag_configure('even',background='gray')
        self.tree.tag_configure('dif',foreground='brown')
        self.tree.tag_configure('work',background='white')
        self.tree.tag_configure('short',background='#F5EFE0')
        self.tree.tag_configure('rest',background='#E0B0B0')
        self.tree.tag_configure('holyday',background='#E7B7A4')
        for c in self.columns:
            self.tree.heading(c,text=c)
            self.tree.column(c,width=65,anchor='center')
        self.tree.column('day',width=30)
        scrX=Scrollbar(f,orient='horizontal',command=self.tree.xview)
        self.tree['xscrollcommand']=scrX.set
        if not SCRY:
            self.scrY=Scrollbar(f,orient='vertical',command=self.yview)
            self.tree['yscrollcommand']=self.scrY.set
        else:
            self.tree['yscrollcommand']=SCRY.set
        self.tree.grid(row=1,column=0,sticky=N+S)
        if not SCRY:
            self.scrY.grid(row=1,column=1,sticky=N+S)
        scrX.grid(row=2,column=0,sticky=E+W)
    def set(self,y,m):
        self.y=y
        self.m=m
        self.show()
    def yview(self,*args):
        self.tree.yview(*args)
        self.yview2(*args)
    def yview2(self,*args):
        pass
    def show(self):
        d=self.data[self.y][self.m]
        V=list(d['degur'].keys())
        self.v['values']=V
        self.v.set(V[0])
        self.select_ver()
    def select_ver(self,*e):
        self.tree.delete(*self.tree.get_children())
        d=self.data[self.y][self.m]
        offset=d['offset']
        v=self.v.get()
        col=[]
        for i,deg in enumerate(d['degurs']):
            self.tree.heading(i+1,text=deg)
            col.append(i+1)
        self.tree.configure(displaycolumns=['day']+col)
        items=dict()

        if 'табель' in d['degur']:
            a=[''.join(x) for x in zip(*[[x for x in d['degur']['план'][j]] \
                    for j in d['degurs']])]
            b=[''.join(x) for x in zip(*[[x for x in d['degur']['табель'][j]] \
                    for j in d['degurs']])]
            c=[x!=y for x,y  in zip(a,b)]
        else:
            c=[False]*32

        for i in range(1,d['days']+1):
            tag = (i+offset) % 7 in [0,6] and 'rest' or 'work'
            if i in d['holydays'] : tag='holyday'
            elif i in d['restdays'] : tag='rest'
            elif i in d['shortdays'] : tag='short'
            elif i in d['workdays'] : tag='work'
            if c[i]: tag=[tag,'dif']
            ii=self.tree.insert('','end',values=['-','-','-','-','-'],tag=tag)
            self.tree.set(ii,column='day',value=i)
            items[i]=ii

        
        for j,s in d['degur'][v].items(): # j-degur
            if not s: continue
            for i,val in enumerate(s[1:-1]):
                if val=='J':
                    val='до'
                elif val=='j':
                    val='од'
                elif val=='a':
                    val='10'
                self.tree.set(items[i+1],column=d['degurs'].index(j)+1,value=val)
            if s[0]=='Н':
                if s[1]=='-':
                    self.tree.set(items[1],column=d['degurs'].index(j)+1,value='Н(8)')
                else:
                    self.tree.set(items[1],column=d['degurs'].index(j)+1,value='!')
            if s[-2]=='Н':
                if s[-1]=='-':
                    self.tree.set(items[len(s)-2],column=d['degurs'].index(j)+1,value='Н(4)')
                else:
                    self.tree.set(items[len(s)-2],column=d['degurs'].index(j)+1,value='!')
        self.calc(self.y,self.m)
    def calc(self,y,m):
        d=self.data[y][m]
        offset=d['offset']
        WH=0
        for i in range(1,d['days']+1):
            if i in d['holydays']: wh=0
            elif i in d['restdays'] : wh=0
            elif i in d['shortdays'] : wh=7
            elif i in d['workdays'] : wh=8
            elif (i+offset) % 7 in [0,6]: wh=0
            else: wh=8
            WH+=wh

class statistic():
    def __init__(self,data,y,v='план'):
        
        w=Toplevel()
        w.wm_title('Статистика за {0} год ({1})'.format(y,v))
        w.columnconfigure(0,weight=1)
        w.rowconfigure(0,weight=1)
        
        cols=data[y]['degurs']  # ЗАГЛУШКА : список дежурных
        
        self.t=Treeview(w,columns=cols)
        self.t.column('#0',width=120)
        for c in cols:
            self.t.heading(c,text=c)
            self.t.column(c,width=65,anchor='center')
        self.t.tag_configure('табель',background='green')
        self.t.tag_configure('ош',background='red')
        self.scrX=Scrollbar(w,orient='horizontal',command=self.t.xview)
        self.scrY=Scrollbar(w,orient='vertical',command=self.t.yview)
        self.t['xscrollcommand']=self.scrX.set
        self.t['yscrollcommand']=self.scrY.set
        self.t.grid(row=0,column=0,sticky=N+S+E+W)
        self.scrX.grid(row=1,column=0,sticky=E+W)
        self.scrY.grid(row=0,column=1,sticky=N+S)
        r=self.t.insert('','end',text='рабочих')
        w=self.t.insert('','end',text='отработано')
        e=self.t.insert('','end',text='дополнительные')
        n=self.t.insert('','end',text='ночные')
        h=self.t.insert('','end',text='праздничные')
        x=self.t.insert('','end',text='xxx')
        rz_root=self.t.insert('','end',text='резерв')
        
        for m in ['01','02','03','04','05','06','07','08','09','10','11','12']:
            d0=data[y]
            if m not in d0: continue
            d=d0[m]
            rez=dict()
            tag=''
            if v=='авто':
                if 'табель' in d['degur']:
                    vv='табель'
                    tag=vv
                else:
                    vv='план'
            elif v=='табель':
                if 'табель' not in d['degur']:
                    vv='план'
                    tag='ош'
                else:
                    vv=v
                    tag=vv
            else:
                vv=v

            for j,s in d['degur'][vv].items():
                rez[j]=analyse2(s,d)
            NUL=(0,0,0,0,0,0,0)
            ww=[rez.get(j,NUL)[0] for j in cols]
            ee=[rez.get(j,NUL)[0]-rez.get(j,NUL)[3]+rez.get(j,NUL)[4] for j in cols]
            xx=[rez.get(j,NUL)[0]-rez.get(j,NUL)[3] for j in cols]
            nn=[rez.get(j,NUL)[1] for j in cols]
            hh=[rez.get(j,NUL)[2] for j in cols]
            rr=[rez.get(j,NUL)[3]-rez.get(j,NUL)[4] for j in cols]
            rz=[rez.get(j,NUL)[5] for j in cols]
            self.t.insert(w,'end',text=m,values=ww,tag=tag)
            self.t.insert(n,'end',text=m,values=nn,tag=tag)
            self.t.insert(h,'end',text=m,values=hh,tag=tag)
            self.t.insert(e,'end',text=m,values=ee,tag=tag)
            self.t.insert(x,'end',text=m,values=xx,tag=tag)
            self.t.insert(r,'end',text=m,values=rr,tag=tag)
            self.t.insert(rz_root,'end',text=m,values=rz,tag=tag)

class statistic_xx():
    def __init__(self,data,y,v='план'):
        
        w=Toplevel()
        w.wm_title('Итоги по месяцам за {0} год ({1}) '.format(y,v))
        w.columnconfigure(0,weight=1)
        w.rowconfigure(0,weight=1)
        
        cols=data[y]['degurs']  # ЗАГЛУШКА : список дежурных
        
        self.t=Treeview(w,columns=cols)
        self.t.column('#0',width=120)
        for c in cols:
            self.t.heading(c,text=c)
            self.t.column(c,width=65,anchor='center')
        self.t.tag_configure('табель',background='green')
        self.t.tag_configure('ош',background='red')
        self.scrX=Scrollbar(w,orient='horizontal',command=self.t.xview)
        self.scrY=Scrollbar(w,orient='vertical',command=self.t.yview)
        self.t['xscrollcommand']=self.scrX.set
        self.t['yscrollcommand']=self.scrY.set
        self.t.grid(row=0,column=0,sticky=N+S+E+W)
        self.scrX.grid(row=1,column=0,sticky=E+W)
        self.scrY.grid(row=0,column=1,sticky=N+S)
        roots=dict()
        for m in ['01','02','03','04','05','06','07','08','09','10','11','12']:
            d0=data[y]
            if m not in d0: continue
            roots[m]=self.t.insert('','end',text=m+' ('+data[y][m]['month']+')')
        #r=self.t.insert('','end',text='рабочих')
        #x=self.t.insert('','end',text='xxx')
        #w=self.t.insert('','end',text='отработано')
        #e=self.t.insert('','end',text='дополнительные')
        #n=self.t.insert('','end',text='ночные')
        #h=self.t.insert('','end',text='праздничные')
        #rz_root=self.t.insert('','end',text='резерв')
        
        for m in ['01','02','03','04','05','06','07','08','09','10','11','12']:
            d0=data[y]
            if m not in d0: continue
            d=d0[m]

            rez=dict()
            tag=''
            if v=='авто':
                if 'табель' in d['degur']:
                    vv='табель'
                    tag=vv
                else:
                    vv='план'
            elif v=='табель':
                if 'табель' not in d['degur']:
                    vv='план'
                    tag='ош'
                else:
                    vv=v
                    tag=vv
            else:
                vv=v

            for j,s in d['degur'][vv].items():
                rez[j]=analyse2(s,d)
            NUL=(0,0,0,0,0,0,0)
            ww=[rez.get(j,NUL)[0] for j in cols]
            ee=[rez.get(j,NUL)[0]-rez.get(j,NUL)[3]+rez.get(j,NUL)[4] for j in cols]
            xx=[rez.get(j,NUL)[0]-rez.get(j,NUL)[3] for j in cols]
            nn=[rez.get(j,NUL)[1] for j in cols]
            hh=[rez.get(j,NUL)[2] for j in cols]
            rr=[rez.get(j,NUL)[3]-rez.get(j,NUL)[4] for j in cols]
            rz=[rez.get(j,NUL)[5] for j in cols]
            self.t.insert(roots[m],'end',text='отработано',values=ww,tag=tag)
            self.t.insert(roots[m],'end',text='рабочие',values=rr,tag=tag)
            self.t.insert(roots[m],'end',text='дополнительные',values=ee,tag=tag)
            self.t.insert(roots[m],'end',text='ночные',values=nn,tag=tag)
            self.t.insert(roots[m],'end',text='праздничные',values=hh,tag=tag)

class statistic_q():
    def __init__(self,data,y,v='план'):
        
        w=Toplevel()
        w.wm_title('Поквартальная статистика за {0} год ({1})'.format(y,v))
        w.columnconfigure(0,weight=1)
        w.rowconfigure(0,weight=1)
        
        cols=data[y]['degurs']  # ЗАГЛУШКА : список дежурных
        
        self.t=Treeview(w,columns=cols)
        self.t.column('#0',width=120)
        for c in cols:
            self.t.heading(c,text=c)
            self.t.column(c,width=65,anchor='center')
        self.t.tag_configure('табель',background='green')
        self.t.tag_configure('ош',background='red')
        self.scrX=Scrollbar(w,orient='horizontal',command=self.t.xview)
        self.scrY=Scrollbar(w,orient='vertical',command=self.t.yview)
        self.t['xscrollcommand']=self.scrX.set
        self.t['yscrollcommand']=self.scrY.set
        self.t.grid(row=0,column=0,sticky=N+S+E+W)
        self.scrX.grid(row=1,column=0,sticky=E+W)
        self.scrY.grid(row=0,column=1,sticky=N+S)
        r=self.t.insert('','end',text='рабочих')
        w=self.t.insert('','end',text='отработано')
        e=self.t.insert('','end',text='дополнительные')
        n=self.t.insert('','end',text='ночные')
        h=self.t.insert('','end',text='праздничные')
        x=self.t.insert('','end',text='xxx')
        rz_root=self.t.insert('','end',text='резерв')
        
        rez=dict()
        wwY=[]
        rrY=[]
        eeY=[]
        xxY=[]
        nnY=[]
        hhY=[]
        rzY=[]
        for mm in [1,2,3,4]:
            mmm=[str((mm-1)*3+x).zfill(2) for x in [1,2,3]]
            mmm=[x for x in mmm if x in data[y]]
            tag=''
            
            k=['табель' in data[y][m]['degur'] for m in mmm]
            #print(k)
            if v=='авто':
                if k==[True, True, True]:
                    vv='табель'
                    tag=vv
                else:
                    vv='план'
            elif v=='табель':
                if k!=[True, True, True]:
                    vv='план'
                    tag='ош'
                else:
                    vv=v
                    tag=vv
            else:
                vv=v
            
            ww=[]
            rr=[]
            ee=[]
            xx=[]
            nn=[]
            hh=[]
            rz=[]
            for m in mmm:
                d=data[y][m]
                for j in cols:
                    s=d['degur'][vv].get(j,'*ООООООООООООООООООООООООООООООО*')
                    rez[j]=analyse2(s,d)
                NUL=(0,0,0,0,0,0,0)
                ww.append([rez.get(j,NUL)[0] for j in cols])
                ee.append([rez.get(j,NUL)[0]-rez.get(j,NUL)[3] + \
                        rez.get(j,NUL)[4] for j in cols])
                xx.append([rez.get(j,NUL)[0]-rez.get(j,NUL)[3] for j in cols])
                rr.append([rez.get(j,NUL)[3]-rez.get(j,NUL)[4] for j in cols])
                nn.append([rez.get(j,NUL)[1] for j in cols])
                hh.append([rez.get(j,NUL)[2] for j in cols])
                rz.append([rez.get(j,NUL)[5] for j in cols])
            ww=[sum(x) for x in zip(*ww)]
            rr=[sum(x) for x in zip(*rr)]
            ee=[sum(x) for x in zip(*ee)]
            xx=[sum(x) for x in zip(*xx)]
            nn=[sum(x) for x in zip(*nn)]
            hh=[sum(x) for x in zip(*hh)]
            rz=[sum(x) for x in zip(*rz)]
            wwY.append(ww)
            rrY.append(rr)
            eeY.append(ee)
            xxY.append(xx)
            nnY.append(nn)
            hhY.append(hh)
            rzY.append(rz)
            self.t.insert(w,'end',text=mm,values=ww,tag=tag)
            self.t.insert(r,'end',text=mm,values=rr,tag=tag)
            self.t.insert(n,'end',text=mm,values=nn,tag=tag)
            self.t.insert(h,'end',text=mm,values=hh,tag=tag)
            self.t.insert(e,'end',text=mm,values=ee,tag=tag)
            self.t.insert(x,'end',text=mm,values=xx,tag=tag)
            self.t.insert(rz_root,'end',text=mm,values=rz,tag=tag)
        wwY=[sum(x) for x in zip(*wwY)]
        rrY=[sum(x) for x in zip(*rrY)]
        eeY=[sum(x) for x in zip(*eeY)]
        xxY=[sum(x) for x in zip(*xxY)]
        nnY=[sum(x) for x in zip(*nnY)]
        hhY=[sum(x) for x in zip(*hhY)]
        rzY=[sum(x) for x in zip(*rzY)]
        self.t.insert(w,'end',text='итого',values=wwY,tag='Y')
        self.t.insert(r,'end',text='итого',values=rrY,tag='Y')
        self.t.insert(n,'end',text='итого',values=nnY,tag='Y')
        self.t.insert(h,'end',text='итого',values=hhY,tag='Y')
        self.t.insert(e,'end',text='итого',values=eeY,tag='Y')
        self.t.insert(x,'end',text='итого',values=xxY,tag='Y')
        self.t.insert(rz_root,'end',text='итого',values=rzY,tag='Y')
            
class otp():
    def __init__(self,data,y,v='план'):
        
        w=Toplevel()
        w.wm_title('Отпуска')
        w.columnconfigure(0,weight=1)
        w.rowconfigure(0,weight=1)
        
        years=sorted(data.keys())
        cols=data['2015']['degurs']  # ЗАГЛУШКА : список дежурных

        self.t=Treeview(w,columns=cols)
        for c in cols:
            self.t.heading(c,text=c)
            self.t.column(c,width=65,anchor='center')
        self.t.tag_configure('табель',background='green')
        self.t.tag_configure('ош',background='red')
        self.scrX=Scrollbar(w,orient='horizontal',command=self.t.xview)
        self.scrY=Scrollbar(w,orient='vertical',command=self.t.yview)
        self.t['xscrollcommand']=self.scrX.set
        self.t['yscrollcommand']=self.scrY.set
        self.t.grid(row=0,column=0,sticky=N+S+E+W)
        self.scrX.grid(row=1,column=0,sticky=E+W)
        self.scrY.grid(row=0,column=1,sticky=N+S)

        for y in years:
            x=self.t.insert('','end',text=y)
            xxY=[]
            for m in ['01','02','03','04','05','06','07','08','09','10','11','12']:
                d0=data[y]
                if m not in d0: continue
                d=d0[m]
                rez=dict()
                tag=''
                if v=='авто':
                    if 'табель' in d['degur']:
                        vv='табель'
                        tag=vv
                    else:
                        vv='план'
                elif v=='табель':
                    if 'табель' not in d['degur']:
                        vv='план'
                        tag='ош'
                    else:
                        vv=v
                        tag=vv
                else:
                    vv=v
                for j,s in d['degur'][vv].items():
                    rez[j]=analyse2(s,d)
                NUL=(0,0,0,0,0,0,0)
                xx=[rez.get(j,NUL)[6] for j in cols]
                xxY.append(xx)
                self.t.insert(x,'end',text=m,values=[x or '-' for x in xx],tag=tag)
            xxY=[sum(x) for x in zip(*xxY)]
            self.t.insert(x,'end',text='итого',values=xxY,tag='Y')


class per():
    def __init__(self,data,y,v='план'):
        
        w=Toplevel()
        w.wm_title('Доп')
        w.columnconfigure(0,weight=1)
        w.rowconfigure(0,weight=1)
        
        years=sorted(data.keys())
        cols=data['2013']['degurs']  # ЗАГЛУШКА : список дежурных

        self.t=Treeview(w,columns=cols)
        for c in cols:
            self.t.heading(c,text=c)
            self.t.column(c,width=65,anchor='center')
        self.t.tag_configure('табель',background='green')
        self.t.tag_configure('ош',background='red')
        #self.t.tag_configure('табель',background='green')
        self.scrX=Scrollbar(w,orient='horizontal',command=self.t.xview)
        self.scrY=Scrollbar(w,orient='vertical',command=self.t.yview)
        self.t['xscrollcommand']=self.scrX.set
        self.t['yscrollcommand']=self.scrY.set
        self.t.grid(row=0,column=0,sticky=N+S+E+W)
        self.scrX.grid(row=1,column=0,sticky=E+W)
        self.scrY.grid(row=0,column=1,sticky=N+S)

        for y in years:
            x=self.t.insert('','end',text=y)
            eeY=[]
            for m in ['01','02','03','04','05','06','07','08','09','10','11','12']:
                d0=data[y]
                if m not in d0: continue
                d=d0[m]
                rez=dict()
                tag=''
                if v=='авто':
                    if 'табель' in d['degur']:
                        vv='табель'
                        tag=vv
                    else:
                        vv='план'
                elif v=='табель':
                    if 'табель' not in d['degur']:
                        vv='план'
                        tag='ош'
                    else:
                        vv=v
                        tag=vv
                else:
                    vv=v
                for j,s in d['degur'][vv].items():
                    rez[j]=analyse2(s,d)
                NUL=(0,0,0,0,0,0,0)
                ee=[rez.get(j,NUL)[0]-rez.get(j,NUL)[3]+rez.get(j,NUL)[4] for j in cols]
                eeY.append(ee)
                self.t.insert(x,'end',text=m,values=[x or '-' for x in ee],tag=tag)
            eeY=[sum(x) for x in zip(*eeY)]
            self.t.insert(x,'end',text='итого',values=eeY,tag='Y')

            
class tabel():
    def __init__(self,data,y,m):
        w=Toplevel()
        w.geometry('{0}x{1}+0+0'.format(120+25*21,w.winfo_screenheight()-80))
        
        w.columnconfigure(0,weight=1)
        w.rowconfigure(0,weight=1)
        
        d=data[y][m]
        v='план'
        if 'табель' in d['degur']:
            v='табель'
        w.wm_title('{0} {1} {2}'.format(v,y,m))
        #deg=data['2013']['11']['degurs']  # ЗАГЛУШКА : список дежурных
        deg=d['degurs']  # ЗАГЛУШКА : список дежурных
        cc=list(range(1,17))
        cols=[str(c) for c in cc]+['s','сум','ноч','пра','доп']
        
        self.t=Treeview(w,columns=cols)
        self.t.column('#0',width=100)
        for c in cols:
            self.t.heading(c,text=c)
            self.t.column(c,width=25,anchor='center')
        self.t.tag_configure('title',background='gray')
        self.scrX=Scrollbar(w,orient='horizontal',command=self.t.xview)
        self.scrY=Scrollbar(w,orient='vertical',command=self.t.yview)
        self.t['xscrollcommand']=self.scrX.set
        self.t['yscrollcommand']=self.scrY.set
        self.t.grid(row=0,column=0,sticky=N+S+E+W)
        self.scrX.grid(row=1,column=0,sticky=E+W)
        self.scrY.grid(row=0,column=1,sticky=N+S)

        rez=dict()
        for j,s in d['degur'][v].items():
                rez[j]=analyse2(s,d)
        for j in deg:
            ww1=[]
            ww2=[]
            a=0
            nn1=[]
            for x in d['degur'][v].get(j,''):
                if a:
                    if x=='Д':
                        ww1.append('!!')
                        nn1+=[0,0]
                        y=a+12
                        a=0
                    elif x in [str(xx) for xx in range(1,10)]:
                        ww1.append('Я')
                        nn1+=[0,0]
                        y=a+int(x)
                        a=0
                    elif x=='Н':
                        ww1.append('!')
                        nn1+=[2,6]
                        y=a+4
                        a=8
                    elif x=='-':
                        ww1.append('Н')
                        nn1+=[0,0]
                        y=a+0
                        a=0
                    else:
                        ww1.append('!')
                        nn1+=[0,0]
                        y=a+0
                        a=0
                else:
                    if x=='Д':
                        ww1.append('Я')
                        nn1+=[0,0]
                        y=12
                        a=0
                    elif x in [str(xx) for xx in range(1,10)]:
                        ww1.append('Я')
                        nn1+=[0,0]
                        y=int(x)
                        a=0
                    elif x=='Н':
                        ww1.append(x)
                        nn1+=[2,6]
                        y=4
                        a=8
                    elif x=='-':
                        ww1.append('В')
                        nn1+=[0,0]
                        y=0
                        a=0
                    else:
                        ww1.append(x)
                        nn1+=[0,0]
                        y=0
                        a=0
                ww2.append(y)
            ww=rez.get(j,(0,0,0,0))[0]
            ee=rez.get(j,(0,0,0,0))[0] -rez.get(j,(0,0,0,0))[3]
            #ee=rez.get(j,(0,0,0,0))[3]
            nn=rez.get(j,(0,0,0,0))[1]
            hh=rez.get(j,(0,0,0,0))[2]
            s1=sum([x and 1 for x in ww2[1:16]])
            s2=sum([x and 1 for x in ww2[16:-1]])
            n0=sum([x=='Н' and 1 or 0 for x in ww1[1:-1]])

            z=self.t.insert('','end',text=j)
            self.t.insert(z,'end',text='',values=list(range(1,16)),tag='title')
            self.t.insert(z,'end',text='',values=ww1[1:16]+['',s1])
            self.t.insert(z,'end',text='',values=[x or '-' for x in ww2[1:16]]+ \
                    ['']+[sum(ww2[1:16])])
            self.t.insert(z,'end',text='',values=list(range(16,32)),tag='title')
            self.t.insert(z,'end',text='',values=ww1[16:-1]+['']*(33-len(ww1))+ \
                    [s2,s1+s2,n0])
            self.t.insert(z,'end',text='',values=[x or '-' for x in ww2[16:-1]]+ \
                    ['']*(16-len(ww2[16:-1]))+[sum(ww2[16:-1]),sum(ww2[1:-1]),sum(nn1[1:-3]),hh,ee])

class Main():
    def __init__(self):
        f=open('degur.yaml','r',encoding='utf-8')
        self.data=yaml.load(f.read())
        f.close()
        root=Tk()
        root.wm_title('Дежурства v 0.1.1 (c) 2013-2015, Shershakov D.')
        root.geometry('{0}x{1}+0+0'.format(root.winfo_screenwidth()-10,root.winfo_screenheight()-80))
        root.rowconfigure(1,weight=1)
        root.columnconfigure(1,weight=1)
        root.columnconfigure(2,weight=1)
        f0=Frame(root)
        f1=Frame(root)
        f2=Frame(root)
        self.y=Combobox(f0,width=4)
        self.m=Combobox(f0,width=4)
        self.y.grid(row=0,column=0)
        self.m.grid(row=1,column=0)
        self.y.bind('<<ComboboxSelected>>',self.setY)
        self.m.bind('<<ComboboxSelected>>',self.setM)
        f0.grid(row=0,column=0,rowspan=10,sticky=N+S+E+W)
        f1.grid(row=1,column=1,sticky=N+S+E+W)
        f2.grid(row=1,column=2,sticky=N+S+E+W)
        self.g1=Gr(f1,self.data)
        self.g2=Gr(f2,self.data,SCRY=self.g1.scrY)
        self.set0()
        self.g1.yview2=self.g2.yview
        root.bind('<F1>',lambda e: MyHelp(root,self.data,self.y.get()))
        root.bind_all('<Control-F3>',lambda e: statistic_q(self.data,self.y.get(),v='план'))
        root.bind_all('<Shift-F3>',lambda e: statistic_q(self.data,self.y.get(),v='табель'))
        root.bind_all('<Control-F4>',lambda e: statistic(self.data,self.y.get(),v='план'))
        root.bind_all('<Shift-F4>',lambda e: statistic(self.data,self.y.get(),v='табель'))
        root.bind_all('<F3>',lambda e: statistic_q(self.data,self.y.get(),v='авто'))
        root.bind_all('<F4>',lambda e: statistic(self.data,self.y.get(),v='авто'))
        root.bind_all('<F5>',lambda e: tabel(self.data,self.y.get(),self.m.get()))
        root.bind_all('<F7>',lambda e: otp(self.data,self.y.get(),v='авто'))
        root.bind_all('<F8>',lambda e: statistic_xx(self.data,self.y.get(),v='авто'))
        root.bind_all('<F9>',lambda e: per(self.data,self.y.get(),v='авто'))
        root.bind_all('<Control-F8>',lambda e: statistic_xx(self.data,self.y.get(),v='план'))
        FreeConsole()
        root.mainloop()
    def set0(self,*e):
        Y=sorted(list(self.data.keys()))
        self.y['values']=Y
        self.y.set(Y[0])
        self.setY(*e)
    def setY(self,*e):
        M=sorted([x for x in self.data[self.y.get()] if str(x).isnumeric()])
        self.m['values']=M
        self.m.set(M[0])
        self.setM(*e)
    def setM(self,*e):
        y=self.y.get()
        m=self.m.get()
        self.g1.set(y,m)
        self.g2.set(y,m)

Main()
