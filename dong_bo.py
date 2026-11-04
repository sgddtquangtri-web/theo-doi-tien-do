#!/usr/bin/env python3
"""Cap nhat data.json tu tep theo doi. Ma tep lay tu bien moi truong DRIVE_FILE_ID."""
import sys,re,json,hashlib,calendar,datetime as dt,glob,os,unicodedata
import docx
import urllib.request
NOW=dt.datetime.now(dt.timezone.utc)
TODAY=(NOW+dt.timedelta(hours=7)).date()
SRC='bc.docx'
if len(sys.argv)>1: SRC=sys.argv[1]
else:
    fid=os.environ['DRIVE_FILE_ID'].strip()
    b=urllib.request.urlopen('https://drive.google.com/uc?export=download&id='+fid,timeout=60).read()
    if b[:2]!=b'PK': sys.exit('Khong tai duoc tep docx: kiem tra quyen chia se cua tep.')
    open(SRC,'wb').write(b)
BO="Bộ GD&ĐT, Trung ương giao";TINH="Tỉnh ủy, UBND tỉnh, sở ngành giao"
def clean(s): return re.sub(r'\s+',' ',unicodedata.normalize('NFC',s).replace('\xa0',' ')).strip()
def dv(s):
    k=clean(s).lower().replace('–','-').replace(' - ','-')
    for key,val in [('kế hoạch','Phòng KHTC'),('tổ chức cán bộ','Phòng TCCB'),('quản lý chất lượng','Phòng QLCL'),('mầm non','Phòng GDMN'),('mần non','Phòng GDMN'),('phổ thông','Phòng GDPT'),('nghề nghiệp','Phòng GDNN-GDTX'),('chính trị','Phòng CTTT-CTHSSV'),('văn phòng','Văn phòng')]:
        if key in k: return val
    return 'Chưa phân công'
def nhom(s):
    k=clean(s).lower()
    return BO if re.search(r'^(bộ |cục )|bộ giáo dục|bộ gd',k) else TINH
def so(s):
    t=clean(s)
    t2=re.sub(r'^(công văn|công báo|kế hoạch|chương trình hành động|chương trình|thông báo|thông tư|cv|kh)\s+','',t,flags=re.I)
    t2=re.sub(r'^(số\s+)?','',t2,flags=re.I)
    m=re.match(r'[\w./&\\-]*\d[\w./&\\()-]*',t2)
    return re.sub(r'\d{1,2}/\d{1,2}/\d{4}$','',m.group(0)).rstrip('.,').replace('\\','') if m else t[:40]
def valid(y,m,d):
    try: return dt.date(y,m,d)
    except ValueError: return None
def nam_toi(d,m):
    x=valid(TODAY.year,m,d)
    return x if x and x>=TODAY else valid(TODAY.year+1,m,d)
def thang_toi(d):
    x=valid(TODAY.year,TODAY.month,min(d,calendar.monthrange(TODAY.year,TODAY.month)[1]))
    if x>=TODAY: return x
    y,m=(TODAY.year+1,1) if TODAY.month==12 else (TODAY.year,TODAY.month+1)
    return valid(y,m,min(d,calendar.monthrange(y,m)[1]))
def han(raw):
    """Tra ve (iso hoac '', co_dinh_ky). Moc dinh ky lay lan den han ke tiep."""
    t=clean(raw).lower()
    if not t or 'không giao' in t or 'thường xuyên' in t: return '',('thường xuyên' in t)
    if 'hàng tháng, quý' in t: return '',True
    m=re.search(r'quý iv/(\d{4})',t)
    if m: return f'{m.group(1)}-12-31',False
    m=re.search(r'trong tháng (\d{1,2})/(\d{4})',t)
    if m:
        y,mo=int(m.group(2)),int(m.group(1));return dt.date(y,mo,calendar.monthrange(y,mo)[1]).isoformat(),False
    if 'tháng cuối quý' in t:
        d=int(re.search(r'ngày\s*(\d{1,2})',t).group(1))
        c=[valid(TODAY.year+a,mo,d) for a in (0,1) for mo in (3,6,9,12)]
        return min(x for x in c if x>=TODAY).isoformat(),True
    m=re.search(r'ngày\s*(\d{1,2})\s*hằng\s*tháng',t)
    if m: return thang_toi(int(m.group(1))).isoformat(),True
    pairs=re.findall(r'(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{4}))?',t)
    if 'hằng năm' in t or 'định kỳ' in t:
        c=[nam_toi(int(a),int(b)) for a,b,_ in pairs]
        c=[x for x in c if x]
        return (min(c).isoformat() if c else ''),True
    m=re.search(r'ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})(?:\s*năm\s*(\d{4}))?',t)
    if m:
        y=int(m.group(3) or 0)
        x=valid(y,int(m.group(2)),int(m.group(1))) if y else nam_toi(int(m.group(1)),int(m.group(2)))
        return (x.isoformat() if x else ''),False
    if pairs:
        a,b,y=pairs[-1]
        y=int(y) if y else TODAY.year
        x=valid(y,int(b),int(a))
        return (x.isoformat() if x else ''),False
    return '',False
def ngay(raw):
    m=re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})',raw)
    x=valid(int(m.group(3)),int(m.group(2)),int(m.group(1))) if m else None
    return x.isoformat() if x else ''
DONE=re.compile(r'đã hoàn thành',re.I)
def tien_do(kq,ly):
    both=f'{kq} {ly}'
    if DONE.search(both): return 'Hoàn thành'
    if re.search(r'^đang',clean(ly),re.I): return 'Đang xử lý'
    if re.search(r'\d+/[\w&-]*SG|tờ trình|báo cáo số|công văn',kq,re.I): return 'Chờ xác nhận'
    if not clean(both) or clean(both).lower()=='chưa có sản phẩm': return 'Chưa xử lý'
    return 'Đang xử lý'
def parse():
    rows=[]
    for r in docx.Document(SRC).tables[0].rows[1:]:
        c=[];seen=[]
        for x in r.cells:
            if x._tc in seen: continue
            seen.append(x._tc);c.append(clean(x.text))
        c+=['']*(8-len(c))
        rows.append(c)
    out={};
    for c in rows:
        nd,bh,chutri,giao,hh,kq,ly=c[1],c[2],c[3],c[4],c[5],c[6],c[7]
        if not nd: continue
        # hang lech cot: o ngay giao chua van ban ket qua
        if len(giao)>12 and not re.match(r'\d',giao):
            kq,ly,giao,hh=giao,(hh+' '+kq).strip(),'',''
        h,dk=han(hh)
        d=dict(nguon='tep',soKyHieu=so(nd),trichYeu=nd,donViBanHanh=bh,nhom=nhom(bh),donViThucHien=dv(chutri),nguoiPhuTrach='',
               han=h,hanGoc=hh,tienDo=tien_do(kq,ly),dinhKy=hh if dk else '',ketQua=kq,ghiChu=ly if not DONE.fullmatch(ly) else '',
               ngayBanHanh='',ngayGiao=ngay(giao),doKhan='')
        base='bc-'+hashlib.sha1(nd.lower().encode()).hexdigest()[:12];i=base;n=1
        while i in out: n+=1;i=f'{base}-{n}'
        out[i]=d
    return out
new=parse()
if len(new)<100: sys.exit(f'Chi doc duoc {len(new)} dong, khong cap nhat.')
SENS=re.compile(r'KPI|đánh giá cán bộ|nhân sự|kỷ luật|kiểm tra, giám sát|khắc phục hạn chế|Thanh tra|Công an|CAT-|PC08|cung cấp thông tin, tài liệu|bồi thường|biên chế|cán bộ|khủng bố|quân sự|phòng không|Kết luận số 213|213/KL',re.I)
pub=[]
for i,x in new.items():
    if SENS.search(' '.join([x['trichYeu'],x['donViBanHanh'],x['ketQua'],x['ghiChu']])): continue
    y={k:x[k] for k in ['soKyHieu','trichYeu','donViBanHanh','nhom','donViThucHien','han','hanGoc','tienDo','dinhKy','ketQua','ngayGiao']}
    y['id']=i;y['nguon']='tep'
    y['ghiChu']=re.sub(r'.*Đ/c.*','',x['ghiChu'])
    y['ketQua']=re.sub(r'https?://\S+','',y['ketQua'])
    pub.append(y)
json.dump({'luc':NOW.isoformat(),'viec':pub},open('data.json','w'),ensure_ascii=False)
print(len(new),'dong,',len(pub),'cong khai')
