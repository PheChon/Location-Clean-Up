import pandas as pd, numpy as np, re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

SRC = '/Users/phachon/Documents/DKSH/location-clean-up/input/Location TH cleanup 29052026.xlsx'
OUT = '/Users/phachon/Documents/DKSH/location-clean-up/output/Location_TH_dedup_step1-3.xlsx'
ACT = "Active SAP confirm by P'Ying 29May2026"

df = pd.read_excel(SRC, sheet_name='Location clean up', dtype=str)
df.columns = [str(c).strip() for c in df.columns]
df = df.rename(columns={'Unnamed: 15': '(helper) clean street spaced'})
n = len(df)

def s(x):
    if x is None: return ''
    x = str(x)
    return '' if x.strip().lower() in ('', 'nan', 'none') else x.strip()

def strip_dot0(x):
    x = s(x)
    return re.sub(r'\.0$', '', x)

# ---------- STEP 1: CLEAN / NORMALIZE ----------
sap   = df['SVMX_SAP_Code__c'].map(strip_dot0)
orgid = df['Organization ID'].map(s)
wo    = pd.to_numeric(df['No. of Work order'].map(strip_dot0), errors='coerce').fillna(0).astype(int)
ib    = pd.to_numeric(df['No. of IB'].map(strip_dot0), errors='coerce').fillna(0).astype(int)
conf  = df[ACT].map(s)
status= df['Status__c'].map(s)
lastmod = pd.to_datetime(df['LastModifiedDate'], errors='coerce', utc=True)

def clean_zip(z):
    z = strip_dot0(z)
    m = re.search(r'\d{5}', z)
    return m.group(0) if m else ''
zip5 = df['SVMXC__Zip__c'].map(clean_zip)

# normalized address key for matching (street + zip), abbreviation-aware
ABBR = ['ถนน','ตำบล','อำเภอ','จังหวัด','แขวง','เขต','หมู่ที่','หมู่',
        'tambon','amphoe','amphur','province','district','subdistrict','sub-district',
        'road','moo']
def norm_street(v):
    t = s(v).lower()
    t = re.sub(r'<br\s*/?>', ' ', t)
    t = t.replace('ถ.',' ').replace('ต.',' ').replace('อ.',' ').replace('จ.',' ').replace('ม.',' ')
    t = t.replace('rd.',' ').replace('rd',' ').replace('t.',' ').replace('a.',' ')
    for w in ABBR:
        t = t.replace(w, ' ')
    t = re.sub(r'[^0-9a-z\u0e00-\u0e7f]', '', t)   # drop punctuation/space, keep thai+latin+digits
    return t
nstreet = df['SVMXC__Street__c'].map(norm_street)
match_key = [(ns + '|' + z) if (len(ns) >= 6 and re.search(r'\d', ns)) else '' for ns, z in zip(nstreet, zip5)]
match_key = pd.Series(match_key, index=df.index)

# ---------- STEP 2: CLUSTER (union-find) ----------
parent = list(range(n))
def find(a):
    while parent[a] != a:
        parent[a] = parent[parent[a]]; a = parent[a]
    return a
def union(a, b):
    ra, rb = find(a), find(b)
    if ra != rb: parent[max(ra,rb)] = min(ra,rb)

# link by SAP code (reliable; may span orgs)
for code, idx in pd.Series(range(n)).groupby(sap.values).groups.items():
    if code and len(idx) > 1:
        idx = list(idx); first = idx[0]
        for j in idx[1:]: union(first, j)
# link by (org id + normalized address), within same org only
ok_addr = (orgid != '') & (match_key != '')
for key, idx in pd.Series(np.arange(n)[ok_addr.values]).groupby(
        (orgid[ok_addr] + '##' + match_key[ok_addr]).values).groups.items():
    if len(idx) > 1:
        idx = list(idx); first = idx[0]
        for j in idx[1:]: union(int(first), int(j))

root = np.array([find(i) for i in range(n)])
df['_root'] = root
csize = df.groupby('_root')['_root'].transform('count')

# ---------- STEP 3: ELECT MASTER + ACTION ----------
status_rank = {'Approved':4,'Draft':3,'Inactive':1,'Duplicate':1,'To be deleted':0,'0':0,'':0}
completeness = sum([(df[c].map(s) != '').astype(int) for c in
    ['SVMXC__Street__c','District__c','SVMXC__City__c','SVMXC__State__c','SVMXC__Zip__c']])
active = (conf == 'Active').astype(int)
txn = wo + ib
recency = lastmod.view('int64').fillna(0) if hasattr(lastmod,'view') else lastmod.astype('int64')
recency = pd.to_numeric(lastmod.astype('int64'), errors='coerce').fillna(0)
srank = status.map(lambda x: status_rank.get(x,0))
locid = df['Location ID'].map(s)

score = pd.DataFrame({'active':active,'txn':txn,'recency':recency,'srank':srank,
                      'comp':completeness,'locid':locid,'root':root})

cluster_id = np.empty(n, dtype=object)
master_yn  = np.array(['']*n, dtype=object)
action     = np.empty(n, dtype=object)
master_loc = np.empty(n, dtype=object)
reason     = np.array(['']*n, dtype=object)
need_rev   = np.array(['']*n, dtype=object)
rev_note   = np.array(['']*n, dtype=object)
method     = np.empty(n, dtype=object)

cid_map = {}; next_cid = 1
for r, grp in score.groupby('root'):
    idx = grp.index.tolist()
    members_sap = sap.iloc[idx]
    members_key = (orgid.iloc[idx] + '##' + match_key.iloc[idx])
    linked_sap  = members_sap[members_sap!=''].duplicated(keep=False).any()
    linked_addr = members_key[match_key.iloc[idx]!=''].duplicated(keep=False).any()
    if len(idx) == 1:
        meth = 'Singleton'
    elif linked_sap and linked_addr: meth = 'SAP+Address'
    elif linked_sap: meth = 'SAP'
    elif linked_addr: meth = 'Address'
    else: meth = 'Mixed'
    cid_map[r] = f'C{next_cid:05d}'; next_cid += 1
    cid = cid_map[r]
    for i in idx:
        cluster_id[i] = cid; method[i] = meth

    if len(idx) == 1:
        i = idx[0]
        master_yn[i]=''; master_loc[i]=locid.iloc[i]; action[i]='Keep (unique)'
        reason[i]='no duplicate found in round 1'
        continue
    # pick master
    g = grp.sort_values(['active','txn','recency','srank','comp','locid'],
                        ascending=[False,False,False,False,False,False])
    mi = g.index[0]
    m_active = active.iloc[mi]==1; m_txn=int(txn.iloc[mi]); m_status=status.iloc[mi]
    for i in idx:
        master_loc[i] = locid.iloc[mi]
        if i == mi:
            master_yn[i]='Y'; action[i]='Master'
            reason[i]=f"{'SAP active' if m_active else 'no active SAP'}; WO+IB={m_txn}; status={m_status or '-'}"
        else:
            master_yn[i]=''; action[i]='Merge'
            reason[i]=f"merge into {locid.iloc[mi]}"
    # review flags
    notes=[]
    n_active = int(active.iloc[idx].sum())
    if n_active>1: notes.append(f'{n_active} active-SAP rows')
    if meth in ('Address',): notes.append('address-only match (verify same site)')
    if orgid.iloc[idx].nunique()>1: notes.append('spans >1 Organization')
    if m_status in ('Inactive','Duplicate','To be deleted','0',''): notes.append(f'master status {m_status or "blank"}')
    holders = int(((wo.iloc[idx]>0)|(ib.iloc[idx]>0)).sum())
    if holders>1: notes.append(f'{holders} rows hold WO/IB (re-parent)')
    if notes:
        for i in idx:
            need_rev[i]='Yes'; rev_note[i]='; '.join(notes)

df['Match key (norm)'] = match_key
df['Cluster ID'] = cluster_id
df['Cluster size'] = csize.values
df['Cluster method'] = method
df['Master (Y/N)'] = master_yn
df['Action'] = action
df['Master Location ID'] = master_loc
df['Master reason'] = reason
df['Needs review'] = need_rev
df['Review note'] = rev_note
df = df.drop(columns=['_root'])

# ---------- SUMMARY ----------
multi = df[df['Cluster size'].astype(int)>1]
summary = {
 'Total location records': n,
 'Distinct clusters': int(df['Cluster ID'].nunique()),
 'Duplicate clusters (size>1)': int(multi['Cluster ID'].nunique()),
 'Records in duplicate clusters': len(multi),
 'Action = Master (survivors of dup groups)': int((df['Action']=='Master').sum()),
 'Action = Merge (fold into master)': int((df['Action']=='Merge').sum()),
 'Action = Keep (unique, recheck in round 2)': int((df['Action']=='Keep (unique)').sum()),
 'Clusters by SAP only': int(multi[multi['Cluster method']=='SAP']['Cluster ID'].nunique()),
 'Clusters by Address only': int(multi[multi['Cluster method']=='Address']['Cluster ID'].nunique()),
 'Clusters by SAP+Address': int(multi[multi['Cluster method']=='SAP+Address']['Cluster ID'].nunique()),
 'Records flagged Needs review': int((df['Needs review']=='Yes').sum()),
}
print('=== STEP 1-3 RESULT ===')
for k,v in summary.items(): print(f'{k:48s}: {v}')

sum_df = pd.DataFrame({'Metric':list(summary.keys()),'Value':list(summary.values())})

# ---------- WRITE + FORMAT ----------
with pd.ExcelWriter(OUT, engine='openpyxl') as xl:
    df.to_excel(xl, sheet_name='Dedup result', index=False)
    sum_df.to_excel(xl, sheet_name='Summary', index=False)

wb = load_workbook(OUT)
hdr_fill = PatternFill('solid', fgColor='1D3A2A')
hdr_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
thin = Side(style='thin', color='D0D0D0')
for ws in wb.worksheets:
    for c in ws[1]:
        c.fill=hdr_fill; c.font=hdr_font; c.alignment=Alignment(vertical='center', wrap_text=True)
    ws.freeze_panes = 'A2'
    ws.row_dimensions[1].height = 30
ws = wb['Dedup result']
ws.auto_filter.ref = ws.dimensions
cols = list(df.columns)
def L(name): return get_column_letter(cols.index(name)+1)
widths = {'Organization name':26,'Location Name':30,'SVMXC__Street__c':26,'Status__c':12,
          'SVMX_SAP_Code__c':14,ACT:16,'Match key (norm)':24,'Cluster ID':11,'Cluster size':6,
          'Cluster method':14,'Master (Y/N)':8,'Action':16,'Master Location ID':20,
          'Master reason':30,'Needs review':8,'Review note':34,'No. of Work order':8,'No. of IB':8}
for nme,w in widths.items():
    if nme in cols: ws.column_dimensions[L(nme)].width=w
last = ws.max_row
arange = f'{L("Action")}2:{L("Action")}{last}'
ws.conditional_formatting.add(arange, CellIsRule(operator='equal', formula=['"Master"'], fill=PatternFill('solid', fgColor='D5EAD9')))
ws.conditional_formatting.add(arange, CellIsRule(operator='equal', formula=['"Merge"'], fill=PatternFill('solid', fgColor='DCE8FB')))
ws.conditional_formatting.add(arange, CellIsRule(operator='equal', formula=['"Keep (unique)"'], fill=PatternFill('solid', fgColor='EFEEE8')))
rrange = f'{L("Needs review")}2:{L("Needs review")}{last}'
ws.conditional_formatting.add(rrange, CellIsRule(operator='equal', formula=['"Yes"'], fill=PatternFill('solid', fgColor='FAE6C8')))
for ws2 in wb.worksheets:
    for row in ws2.iter_rows(min_row=2):
        for c in row:
            if c.font is None or c.font.name!='Arial':
                c.font=Font(name='Arial', size=10)
ws.column_dimensions[L('Summary' if False else 'Action')].width=16
wb['Summary'].column_dimensions['A'].width=46
wb['Summary'].column_dimensions['B'].width=14
wb.save(OUT)
print('\nSaved:', OUT)