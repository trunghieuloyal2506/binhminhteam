"""
Cập nhật timesheet_server.py với:
1. Đơn vị NGÀY thay vì giờ (÷ 8.5)
2. Thêm header CORS đầy đủ hơn
"""
import json, calendar, io
from http.server import HTTPServer, BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date

CN='1F3864'; CL='D9E1F2'; CS='BDD7EE'; CU='FFDDC1'
CWE='D6DCE4'; CY='FFF2CC'; CO='F4B942'; CG='E2EFDA'; CW='FFFFFF'

def F(c): return PatternFill('solid', fgColor=c)
def Fn(bold=False, sz=9, col='000000', italic=False):
    return Font(name='Arial', bold=bold, size=sz, color=col, italic=italic)
def Al(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def Bd():
    t=Side(style='thin')
    return Border(left=t, right=t, top=t, bottom=t)

def S(ws, r, ci, val=None, bold=False, sz=9, fg=None, fc='000000',
      h='center', v='center', wrap=False, bd=True, italic=False, nf=None):
    c=ws.cell(r,ci)
    if val is not None: c.value=val
    c.font=Fn(bold,sz,fc,italic); c.alignment=Al(h,v,wrap)
    if fg: c.fill=F(fg)
    if bd: c.border=Bd()
    if nf: c.number_format=nf
    return c

def M(ws,r1,c1,r2,c2):
    ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)

def hrs_to_days(h, day_hrs=8.5):
    """Convert hours to days, rounded to 2 decimal places"""
    if not h: return 0
    return round(float(h)/day_hrs, 2)

def build_sheet(wb, member, month, year, tasks, leave_map, sheet_name=None):
    days=calendar.monthrange(year,month)[1]
    DAY=['Mon','Tue','Wed','Thu','Fri','Sat','Sun']
    MN=['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

    ws=wb.create_sheet(sheet_name or member['name'][:28])
    TC=5+days

    ws.column_dimensions['A'].width=30; ws.column_dimensions['B'].width=10
    ws.column_dimensions['C'].width=16; ws.column_dimensions['D'].width=7
    for i in range(1,days+1): ws.column_dimensions[get_column_letter(4+i)].width=4.2
    ws.column_dimensions[get_column_letter(TC)].width=7

    ws.row_dimensions[1].height=24
    ws.row_dimensions[8].height=20
    ws.row_dimensions[9].height=16

    # Title
    M(ws,1,1,1,TC); S(ws,1,1,'MONTHLY TIME SHEET',bold=True,sz=14,fg=CN,fc=CW,h='center')
    ws.row_dimensions[2].height=5

    # Info
    wd=sum(1 for dd in range(1,days+1) if date(year,month,dd).weekday()<5)
    est=round(wd*1.0,1)  # days (not hours)
    for r,(lbl,val) in [(3,('Full Name:',member['name'])),
                        (4,('Position:',member.get('role','') or 'QS Engineer')),
                        (5,('Month / Year:',f"{MN[month]} {year}"))]:
        ws.row_dimensions[r].height=17
        S(ws,r,1,lbl,bold=True,sz=10,fg=CL,h='left')
        M(ws,r,2,r,TC); c=ws.cell(r,2); c.value=val
        c.font=Fn(sz=10); c.alignment=Al(h='left'); c.fill=F(CW); c.border=Bd()

    ws.row_dimensions[6].height=17
    S(ws,6,1,'Working days:',bold=True,sz=10,fg=CL,h='left')
    S(ws,6,2,wd,bold=True,sz=10,fg=CW)
    S(ws,6,3,'Estimated Days:',bold=True,sz=10,fg=CL,h='left')
    S(ws,6,4,est,bold=True,sz=10,fg=CW)
    M(ws,6,5,6,TC); ws.cell(6,5).fill=F(CW); ws.cell(6,5).border=Bd()
    ws.row_dimensions[7].height=5

    # Day headers
    M(ws,8,1,9,1); S(ws,8,1,'Brief Description',bold=True,sz=9,fg=CN,fc=CW,h='center',wrap=True)
    M(ws,8,2,9,2); S(ws,8,2,'Code',bold=True,sz=9,fg=CN,fc=CW)
    M(ws,8,3,9,3); S(ws,8,3,'Package',bold=True,sz=9,fg=CN,fc=CW)
    M(ws,8,4,9,4); S(ws,8,4,'Task',bold=True,sz=9,fg=CN,fc=CW)

    for dd in range(1,days+1):
        ci=4+dd; wi=date(year,month,dd).weekday()
        bg=CS if wi==5 else (CU if wi==6 else CN)
        S(ws,8,ci,DAY[wi],bold=True,sz=8,fg=bg,fc=CW)
        S(ws,9,ci,dd,bold=True,sz=8,fg=bg,fc=CW)

    M(ws,8,TC,9,TC); S(ws,8,TC,'Total',bold=True,sz=9,fg=CO,fc=CW)

    # Section grouping
    pre  =[t for t in tasks if t.get('section')=='pre' or t.get('phase') in('cost','pre')]
    post =[t for t in tasks if t.get('section')=='post' or (t.get('phase') in('post','qs') and t not in pre)]
    other=[t for t in tasks if t not in pre and t not in post]
    cur=10

    def sec(row,lbl):
        ws.row_dimensions[row].height=16
        M(ws,row,1,row,TC); c=ws.cell(row,1)
        c.value=lbl; c.font=Fn(True,9,'1F3864',italic=True)
        c.fill=F(CL); c.alignment=Al(h='left'); c.border=Bd()
        return row+1

    def trow(row,t):
        ws.row_dimensions[row].height=15
        hrs=t.get('hours',{})
        S(ws,row,1,t.get('projName',''),sz=9,fg=CG,h='left')
        S(ws,row,2,t.get('projCode',''),sz=9,fg=CW,h='center')
        S(ws,row,3,t.get('package',''),sz=8,fg=CW,h='center')
        S(ws,row,4,t.get('taskCode',''),sz=9,fg=CW,h='center')
        tot=0
        for dd in range(1,days+1):
            ci=4+dd; wi=date(year,month,dd).weekday(); is_we=wi>=5
            hv=float(hrs.get(str(dd),0) or hrs.get(dd,0) or 0)
            dv=hrs_to_days(hv)  # convert to days
            bg=CWE if is_we else CW
            if dv:
                S(ws,row,ci,dv,sz=9,fg=bg,h='center',nf='0.00'); tot+=dv
            else:
                c=ws.cell(row,ci); c.fill=F(bg); c.border=Bd(); c.alignment=Al()
        S(ws,row,TC,round(tot,2) if tot else 0,bold=True,sz=9,fg=CY,h='center',nf='0.00')
        return row+1

    def erow(row):
        ws.row_dimensions[row].height=8
        for ci in range(1,TC+1):
            dd=ci-4; is_we=(ci>4)and(1<=dd<=days)and date(year,month,dd).weekday()>=5
            c=ws.cell(row,ci); c.fill=F(CWE if is_we and ci>4 else CW); c.border=Bd()
        return row+1

    if pre:
        cur=sec(cur,'Pre-Contract works')
        for t in pre: cur=trow(cur,t)
        cur=erow(cur)
    if post:
        cur=sec(cur,'Post-Contract works')
        for t in post: cur=trow(cur,t)
        cur=erow(cur)
    if other:
        cur=sec(cur,'Other works')
        for t in other: cur=trow(cur,t)
        cur=erow(cur)

    # Leave rows (in days)
    for lbl,code in [('Travelling - outside office','TT'),('Sick Leave','SL'),
                     ('Special Leave','SPL'),('Annual Leave','AL'),('Public Holidays','PL')]:
        lhrs=leave_map.get(code,{})
        t={'projName':lbl,'projCode':code,'package':'','taskCode':'',
           'section':'','phase':'',
           'hours':{str(k):v for k,v in lhrs.items()}}
        cur=trow(cur,t)

    # Total per day
    tr=cur; ws.row_dimensions[tr].height=18; ds=10; de=tr-1
    M(ws,tr,1,tr,4); S(ws,tr,1,'Total per day',bold=True,sz=9,fg=CN,fc=CW,h='left')
    for dd in range(1,days+1):
        ci=4+dd; cl=get_column_letter(ci); wi=date(year,month,dd).weekday()
        bg=CWE if wi>=5 else CN
        c=ws.cell(tr,ci); c.value=f'=SUM({cl}{ds}:{cl}{de})'
        c.font=Fn(True,9,CW); c.fill=F(bg); c.alignment=Al(); c.border=Bd(); c.number_format='0.00'
    c=ws.cell(tr,TC); c.value=f'=SUM({get_column_letter(TC)}{ds}:{get_column_letter(TC)}{de})'
    c.font=Fn(True,9,CW); c.fill=F(CO); c.alignment=Al(); c.border=Bd(); c.number_format='0.00'

    # Notes
    nr=tr+2; ws.row_dimensions[nr].height=16; M(ws,nr,1,nr,TC)
    ws.cell(nr,1).value='Notes (Unit: Days)'; ws.cell(nr,1).font=Fn(True,10)
    ws.cell(nr,1).alignment=Al(h='left')
    for i,n in enumerate([
        '1. Unit: Days (1 day = 8.5 hours). E.g. 0.5 = half day, 1.0 = full day.',
        '2. The time recorded against each day (except Saturdays and Sundays) shall equal 1.0 day.',
        '3. Please identify clearly the Pre-Contract or Post Contract Works'],1):
        r=nr+i; ws.row_dimensions[r].height=14; M(ws,r,1,r,TC)
        ws.cell(r,1).value=n; ws.cell(r,1).font=Fn(sz=9); ws.cell(r,1).alignment=Al(h='left')

    # Signatures
    sr=nr+6; ws.row_dimensions[sr].height=18; m1=4; m2=TC//2
    M(ws,sr,1,sr,m1);    S(ws,sr,1,'PREPARED BY',bold=True,sz=9,fg=CL)
    M(ws,sr,m1+1,sr,m2); S(ws,sr,m1+1,'REVIEWED BY',bold=True,sz=9,fg=CL)
    M(ws,sr,m2+1,sr,TC); S(ws,sr,m2+1,'CHECKED BY',bold=True,sz=9,fg=CL)
    name_r=sr+5; ws.row_dimensions[name_r].height=16
    M(ws,name_r,1,name_r,m1);    S(ws,name_r,1,member['name'],bold=True,sz=10,h='center',fg=CW)
    M(ws,name_r,m1+1,name_r,m2); S(ws,name_r,m1+1,'(Line Manager)',sz=9,h='center',fg=CW)
    M(ws,name_r,m2+1,name_r,TC); S(ws,name_r,m2+1,'Võ Thị Duyên - HR Dept',sz=9,h='center',fg=CW)

    ws.freeze_panes='E10'

def build_summary(wb, all_md, month, year):
    MN=['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    ws=wb.create_sheet('SUMMARY')
    ws.column_dimensions['A'].width=5; ws.column_dimensions['B'].width=20
    ws.column_dimensions['C'].width=14; ws.column_dimensions['D'].width=12
    ws.column_dimensions['E'].width=12; ws.column_dimensions['F'].width=12
    ws.column_dimensions['G'].width=14

    M(ws,1,1,1,7); c=ws.cell(1,1)
    c.value=f'TEAM TIMESHEET SUMMARY — {MN[month]} {year} (Unit: Days)'
    c.font=Fn(True,13,CW); c.fill=F(CN); c.alignment=Al(); ws.row_dimensions[1].height=24

    ws.row_dimensions[3].height=18
    for ci,(lbl,h) in enumerate([('STT','center'),('Thành viên','left'),('Chức vụ','left'),
                                   ('Ngày CV','center'),('Ngày nghỉ','center'),
                                   ('Tổng ngày','center'),('% Công suất','center')],1):
        c=ws.cell(3,ci); c.value=lbl; c.font=Fn(True,10,CW)
        c.fill=F(CN); c.alignment=Al(h=h); c.border=Bd()

    total_all=0
    for i,md in enumerate(all_md,1):
        r=3+i; ws.row_dimensions[r].height=16
        work_h=sum(sum(float(v) for v in t.get('hours',{}).values()) for t in md['tasks'])
        leave_h=sum(sum(float(v) for v in lv.values()) for lv in md['leave_map'].values())
        work_d=round(hrs_to_days(work_h),2)
        leave_d=round(hrs_to_days(leave_h),2)
        total_d=round(work_d+leave_d,2)
        cap=md['working_days']
        pct=round(total_d/cap*100,1) if cap>0 else 0
        total_all+=total_d
        for ci,(val,h) in enumerate([(i,'center'),(md['member']['name'],'left'),
                (md['member'].get('role','') or 'QS Engineer','left'),
                (work_d,'center'),(leave_d,'center'),(total_d,'center'),(f'{pct}%','center')],1):
            c=ws.cell(r,ci); c.value=val; c.font=Fn(sz=10)
            c.alignment=Al(h=h); c.border=Bd()
            if i%2==0: c.fill=F('F5F5F5')
        pct_c=ws.cell(r,7)
        pct_c.font=Fn(True,10,'EF4444' if pct>120 else ('F97316' if pct>110 else ('F59E0B' if pct>100 else '15803D')))

    tr=3+len(all_md)+1; ws.row_dimensions[tr].height=18
    M(ws,tr,1,tr,2); S(ws,tr,1,'TỔNG CỘNG',bold=True,sz=11,fg=CY,h='center')
    for ci in [3,4,5]: S(ws,tr,ci,'',fg=CY)
    S(ws,tr,6,round(total_all,2),bold=True,sz=11,fg=CY,h='center',nf='0.00')
    S(ws,tr,7,'',fg=CY)
    ws.freeze_panes='A4'

class Handler(BaseHTTPRequestHandler):
    def log_message(self,format,*args):
        print(f"[Server] {args[0]} {args[1]} {args[2]}")

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors(); self.end_headers()

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Methods','POST,OPTIONS,GET')
        self.send_header('Access-Control-Allow-Headers','Content-Type,Accept')

    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type','text/plain')
        self._cors(); self.end_headers()
        self.wfile.write(b'VinaQS Timesheet Server OK')

    def do_POST(self):
        try:
            length=int(self.headers.get('Content-Length',0))
            body=self.rfile.read(length)
            data=json.loads(body)
            month=int(data['month']); year=int(data['year'])
            members_data=data['members']
            days_in_month=calendar.monthrange(year,month)[1]
            wd=sum(1 for dd in range(1,days_in_month+1) if date(year,month,dd).weekday()<5)

            wb=Workbook(); wb.remove(wb.active)
            all_md=[]
            for md in members_data:
                md['working_days']=wd; all_md.append(md)
                name=md['member']['name'][:28].replace('/','_')
                build_sheet(wb,md['member'],month,year,md['tasks'],md.get('leave_map',{}),name)
            build_summary(wb,all_md,month,year)

            buf=io.BytesIO(); wb.save(buf); buf.seek(0); xlsx=buf.read()
            self.send_response(200)
            self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition',f'attachment; filename="TimeSheet_{month:02d}_{year}_Team.xlsx"')
            self.send_header('Content-Length',str(len(xlsx)))
            self._cors(); self.end_headers()
            self.wfile.write(xlsx)
            print(f"[Server] ✅ Exported {len(members_data)} members")
        except Exception as e:
            import traceback; traceback.print_exc()
            self.send_response(500); self.send_header('Content-Type','application/json')
            self._cors(); self.end_headers()
            self.wfile.write(json.dumps({'error':str(e)}).encode())

if __name__=='__main__':
    print("="*50)
    print("  VinaQS Timesheet Server (Unit: Days)")
    print("  http://localhost:5000")
    print("  Để dừng: Ctrl+C")
    print("="*50)
    HTTPServer(('localhost',5000),Handler).serve_forever()
