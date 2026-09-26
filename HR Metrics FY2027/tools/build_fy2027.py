"""
Build the FY2027 HR Metrics workbook set from the GCG (33 Companies) FY2027 model.

Chain:
  Division workbooks (Auto, Building, Head Office, Manufacturing, Services, Shipping)
      monthly tabs -> Quarter -> GEL HR Metrics FY2027
  GCG cluster workbooks (HDQ, Central America, North Caribbean, South America, South Caribbean)
      monthly tabs -> 0 GCG (33 Companies) HR Metrics FY2027 (monthly tabs) -> Quarter -> GEL HR Metrics FY2027
  Company tabs (one per company) compute monthly Absence Rates and Time-to-Hire YTD that feed the monthly tabs.
"""
import copy, datetime, math, os, re, sys, urllib.parse
import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.formula.tokenizer import Tokenizer, Token
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.external_link.external import ExternalLink, ExternalBook, ExternalSheetNames
from openpyxl.packaging.relationship import Relationship
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from xlhash import new_hash

SRC = os.path.join(HERE, 'src')
MODEL = os.path.join(SRC, 'model.xlsx')
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'out')
PASSWORD = 'GCG'
SALT, HASH = new_hash(PASSWORD)

MONTHS = ['OCT', 'NOV', 'DEC', 'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP']
MONTH_DATES = [datetime.datetime(2026 + (1 if i >= 3 else 0), (10 + i - 1) % 12 + 1, 1) for i in range(12)]
Q_MONTH_COL = dict(zip(MONTHS, ['C', 'D', 'E', 'G', 'H', 'I', 'K', 'L', 'M', 'O', 'P', 'Q']))
Q_COLS = ['F', 'J', 'N', 'R', 'S']          # Q1..Q4, FY in Quarter tab
Q_LABELS = ['Q1', 'Q2', 'Q3', 'Q4', 'FY2027']
FIRST_ROW, LAST_ROW = 3, 70

# ---------------------------------------------------------------- model ----
model = openpyxl.load_workbook(MODEL)
M_NOV = model['NOV']      # clean month template (OCT holds test data)
M_OCT = model['OCT']
M_Q = model['Quarter']

def ftext(v):
    return v.text if isinstance(v, ArrayFormula) else v

# Input rows = rows whose company cell (column C) has no formula in the model month tab
INPUT_ROWS = [r for r in range(FIRST_ROW, LAST_ROW + 1)
              if M_NOV.cell(r, 3).value is None and r != 31]
TAB_ROWS = {39: 'AB', 40: 'AD', 46: 'G', 47: 'O'}        # monthly rows fed by the company tab
AGG = {r: M_NOV[f'AK{r}'].value for r in range(FIRST_ROW, LAST_ROW + 1)}

# ------------------------------------------------------------ helpers -----
CELL_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")

def remap_cols(formula, colmap):
    """Replace column letters inside range operands (sheet prefix untouched)."""
    tok = Tokenizer(formula)
    for t in tok.items:
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            v = t.value
            pre, sep, ref = v.rpartition('!')
            ref = CELL_RE.sub(lambda m: m.group(1) + colmap.get(m.group(2), m.group(2)) + m.group(3) + m.group(4), ref)
            t.value = pre + sep + ref
    return tok.render()

def translate(formula, src, dst):
    return Translator(formula, origin=src).translate_formula(dst)

def cstyle(src, dst):
    dst.font = copy.copy(src.font)
    dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border)
    dst.alignment = copy.copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy.copy(src.protection)

def unlock(cell):
    p = copy.copy(cell.protection); p.locked = False; cell.protection = p

def lock(cell):
    p = copy.copy(cell.protection); p.locked = True; cell.protection = p

def protect(ws, rows_editable=False):
    ws.protection = SheetProtection(sheet=True, objects=True, scenarios=True,
                                    algorithmName='SHA-512', hashValue=HASH, saltValue=SALT, spinCount=100000,
                                    formatColumns=False, formatRows=False,
                                    insertRows=not rows_editable, deleteRows=not rows_editable)

def q(sheet):
    return "'" + sheet.replace("'", "''") + "'" if re.search(r"[^A-Za-z0-9_]", sheet) else sheet

def qnf(r):
    nf = M_Q[f'F{r}'].number_format
    return '#,##0.0' if nf == 'General' else nf

def blank_if(ref):
    return f'=IF({ref}="","",{ref})'

def add_ext_link(wb, rel_target, sheet_names):
    book = ExternalBook(sheetNames=ExternalSheetNames(sheetName=list(sheet_names)), id='rId1')
    link = ExternalLink(externalBook=book)
    link.file_link = Relationship(type='externalLinkPath', Target=rel_target, TargetMode='External', Id='rId1')
    wb._external_links.append(link)
    return len(wb._external_links)

def rel_url(from_dir, to_path):
    return urllib.parse.quote(os.path.relpath(to_path, from_dir).replace(os.sep, '/'), safe="/()'&,")

# ------------------------------------------------ fixed model formulas -----
# (the model is kept 1:1 except where a formula can error or points at the wrong cell)
def fixed_company_formula(r, col):
    """Formula for a derived monthly row in company/total column `col` (row r)."""
    f = {
        11: '=IF(COUNT({c}6:{c}9)=0,"",N({c}6)+N({c}7)*0.75+N({c}8)*0.5+N({c}9)*0.25)',
        25: '=IF(COUNT({c}31,{c}32)<2,"",IF(({c}32+{c}31)=0,"",N({c}22)/(({c}32+{c}31)/2)))',
        26: '=IF(COUNT({c}31,{c}32)<2,"",IF(({c}32+{c}31)=0,"",N({c}23)/(({c}32+{c}31)/2)))',
        29: '=IF(COUNT({c}31,{c}32)<2,"",IF(({c}32+{c}31)=0,"",N({c}27)/(({c}32+{c}31)/2)))',
        30: '=IF(COUNT({c}31,{c}32)<2,"",IF(({c}32+{c}31)=0,"",N({c}28)/(({c}32+{c}31)/2)))',
        33: '=IF(COUNT({c}31,{c}32)<2,"",IF(({c}32+{c}31)=0,"",N({c}24)/(({c}32+{c}31)/2)))',
        37: '=IF(COUNT({c}5,{c}36)=0,"",N({c}5)+N({c}36))',
        68: '=IF(COUNT({c}58:{c}59)=0,"",IFERROR((N({c}58)+N({c}59))/{c}5,""))',
        69: '=IF(COUNT({c}58:{c}59)=0,"",IFERROR((N({c}58)+N({c}59))/{c}56,""))',
    }.get(r)
    return f.format(c=col) if f else None

# ----------------------------------------------------------- specs --------
FY26 = {
    'HDQ': 'GCG HR Metrics/GCG HDQ HR Metrics FY2026.xlsx',
    'CAM': 'GCG HR Metrics/GCG Central America HR Metrics FY2026.xlsx',
    'NCA': 'GCG HR Metrics/GCG North Caribbean HR Metrics FY2026.xlsx',
    'SAM': 'GCG HR Metrics/GCG South America HR Metrics FY2026.xlsx',
    'SCA': 'GCG HR Metrics/GCG South Caribbean HR Metrics FY2026.xlsx',
}
# model column -> (cluster, FY2027 company tab, FY2026 tab used to size the Time-to-Hire table)
GCG_MAP = {
    'C': ('HDQ', '(GCG_HDQ)', '(GCG_HDQ)'),
    'D': ('CAM', '(Culinary_CRI)', 'Culinary_CRI'), 'E': ('CAM', '(Food_GTM)', '(Food_GTM)'),
    'F': ('CAM', '(GCG_GTM)', '(GCG_GTM)'), 'G': ('CAM', '(Ground_HND)', '(Ground_HND)'),
    'H': ('CAM', '(GCG_HND)', '(GCG_HND)'), 'I': ('CAM', '(Concess_PAN)', '(Concess_PAN)'),
    'J': ('CAM', '(GCG_SLV)', '(GCG_SLV)'), 'K': ('CAM', '(Food_SLV)', '(Food_SLV)'),
    'L': ('NCA', '(GCG_ATG)', '(GCG_ATG)'), 'M': ('NCA', '(GCG_GCM)', '(GCG_GCM)'),
    'N': ('NCA', '(Ground_JAM)', '(Ground_JAM)'), 'O': ('NCA', '(GCG_JAM)', '(GCG_JAM)'),
    'P': ('NCA', '(Ground_LCA)', '(Ground_LCA)'), 'Q': ('NCA', '(GCG_LCA)', '(GCG_LCA)'),
    'R': ('NCA', '(Ground_STT)', '(Ground_STT)'),
    'S': ('SAM', '(Bogota)', '(Bogota)'), 'T': ('SAM', '(Colombia)', '(Colombia)'),
    'U': ('SAM', '(Guayaquil)', '(Guayaquil)'), 'V': ('SAM', '(Paraguay)', '(Paraguay)'),
    'W': ('SAM', '(Quito)', '(Quito)'), 'X': ('SAM', '(Uruguay)', '(Uruguay)'), 'Y': ('SAM', '(Caracas)', '(Caracas)'),
    'Z': ('SCA', '(Calloway_AUA)', '(Calloway_AUA)'), 'AA': ('SCA', '(Airport_BRB)', '(Airport_BRB)'),
    'AB': ('SCA', '(GCG_BRB)', '(GCG_BRB)'), 'AC': ('SCA', '(Ground_BRB)', '(Ground_BRB)'),
    'AD': ('SCA', '(GCG_CUR)', '(GCG_CUR)'), 'AE': ('SCA', '(Events_CUR)', '(Events_CUR)'),
    'AF': ('SCA', '(GCG_SXM_Sky)', '(GCG_SXM_Sky)'), 'AG': ('SCA', '(GCG_SXM_Airport)', '(GCG_SXM_Airport)'),
    'AH': ('SCA', '(Allied_TTO)', '(Allied_TTO)'), 'AI': ('SCA', '(Katerserv_TTO)', '(Katerserv_TTO)'),
}
CLUSTERS = [  # key, title, output file
    ('HDQ', 'GCG HDQ', 'GCG HDQ HR Metrics FY2027.xlsx'),
    ('CAM', 'GCG CENTRAL AMERICA', 'GCG Central America HR Metrics FY2027.xlsx'),
    ('NCA', 'GCG NORTH CARIBBEAN', 'GCG North Caribbean HR Metrics FY2027.xlsx'),
    ('SAM', 'GCG SOUTH AMERICA', 'GCG South America HR Metrics FY2027.xlsx'),
    ('SCA', 'GCG SOUTH CARIBBEAN', 'GCG South Caribbean HR Metrics FY2027.xlsx'),
]
GCG_DIR = 'GCG HR Metrics/FY 2027'
GCG_FILE = '0 GCG (33 Companies) HR Metrics FY2027.xlsx'

DIVISIONS = [  # title, folder, output file, FY2026 file, [(name, FY2027 tab, FY2026 tab)]
    ('AUTOMOTIVE', 'Auto HR Metrics', 'Auto HR Metrics FY2027.xlsx', 'Auto HR Metrics/Auto HR Metrics FY2026.xlsx', [
        ('Courtesy Garage Limited', '(Courtesy)', '(Courtesy)'),
        ('Fidelity Motors Limited', '(Fidelity)', '(Fidelity)'),
        ("Peter's Holdings Limited", '(Peter´s)', '(Peter´s)'),
        ('Coreas Hazells (Auto)', '(Coreas_Auto)', '(Coreas_Auto)'),
        ('Jonas Browne & Hubbard (Auto)', '(JBH_Auto)', 'JBH_Auto)')]),
    ('BUILDING SUPPLIES', 'Building Supplies HR Metrics', 'Building HR Metrics FY2027.xlsx', 'Building Supplies HR Metrics/Building HR Metrics FY2026.xlsx', [
        ('Antoseptic Ltd, (BRB)', '(AntoSept)', '(AntoSept)'),
        ('Coreas Hazells Incorporated', '(Coreas)', '(Coreas)'),
        ('M&C Home Depot LTD (SLU)', '(M&C)', '(M&C)'),
        ('Marshall Trading Ltd, (BRB)', '(Marshall)', '(Marshall)'),
        ('Jonas Browne & Hubbard (Building)', '(JBH_Build)', '(JBH_Build)')]),
    ('HEAD OFFICE', 'Head Office HR Metrics', 'Head Office HR Metrics FY2027.xlsx', 'Head Office HR Metrics/Head Office HR Metrics FY2026.xlsx', [
        ('Goddard Enterprises Limited (Barbados)', '(GODDARD)', '(GODDARD)')]),
    ('MANUFACTURING', 'Manufacturing HR Metrics', 'Manufacturing HR Metrics FY2027.xlsx', 'Manufacturing HR Metrics/Manufacturing HR Metrics FY2026.xlsx', [
        ('GEL Manufacturing Holding Co. Ltd.', '(GEL_Holding)', 'GEL_Holding'),
        ('CLC Domicana (Dominican Republic)', '(Dom_CLC)', '(Dom_CLC)'),
        ('Caribbean Label Crafts Ltd. (Barbados)', '(Car_Label)', '(Car_Label)'),
        ('Ecuakao (Ecuador)', '(Ecuakao)', '(Ecuakao)'),
        ('Hipac Ltd. (Barbados)', '(Hipac)', '(Hipac)'),
        ('Label Craft Jamaica Ltd. (Jamaica)', '(Jam_Label)', '(Jam_Label)'),
        ('McBride (Caribbean) Ltd.', '(Mc_Bride)', '(Mc_Bride)'),
        ('Precision Packaging Inc. (Barbados)', '(Precision)', '(Precision)'),
        ('Purity Bakeries Ltd. (Barbados)', '(Purity)', '(Purity)')]),
    ('SERVICES', 'Services HR Metrics', 'Services HR Metrics FY2027.xlsx', 'Services HR Metrics/Services HR Metrics FY2026.xlsx', [
        ("Jonas Browne & Hubbard (G'da) Limited", '(Jonas)', '(Jonas)'),
        ('M&C Drugstore', '(M&C_Drug)', '(M&C_Drug)'),
        ('M&C Ltd.', '(M&C_Ltd)', '(M&C_Ltd)'),
        ('Coreas Hazells (Services)', '(Coreas_Serv)', '(Coreas_Serv)')]),
    ('SHIPPING', 'Shipping HR Metrics', 'Shipping HR Metrics FY2027.xlsx', 'Shipping HR Metrics/Shipping HR Metrics FY2026.xlsx', [
        ('M&C Shipping', '(M&C_Shipp)', '(M&C_Shipp)'),
        ('Coreas Hazells (Shipping)', '(Coreas_Shipp)', '(Coreas_Shipp)'),
        ('Jonas Browne & Hubbard (Shipping)', '(JBH_Shipp)', '(JBH_Shipp)'),
        ('Goddards Shipping Barbados', '(Shipping_BRB)', '(Shipping_BRB)')]),
]
GEL_DIR = 'All GEL Divisions FY2025'   # same SharePoint folder as GEL HR Metrics FY2026
GEL_FILE = 'GEL HR Metrics FY2027.xlsx'

_fy26_cache = {}
def tth_rows(fy26_file, fy26_tab):
    """Rows for the Time-to-Hire tables: FY2026 usage + 25 % head-room, min 40, rounded to 10."""
    if fy26_file not in _fy26_cache:
        _fy26_cache[fy26_file] = openpyxl.load_workbook(os.path.join(SRC, fy26_file))
    ws = _fy26_cache[fy26_file][fy26_tab]
    used = 0
    for r in range(21, ws.max_row + 1):
        if any(ws.cell(r, c).value not in (None, '') for c in (2, 4, 5, 9, 11, 12)):
            used = r - 20
    return max(40, int(math.ceil(used * 1.25 / 10.0)) * 10)

# ------------------------------------------------------- monthly tabs -----
def build_month(ws, mi, title, companies, total_col, feed):
    """companies: list of dicts {col, code, name, tab}; feed(r, col, month) -> formula or None for inputs"""
    month = MONTHS[mi]
    last_col = companies[-1]['col']
    agg_col, chk_col = L(CI(total_col) + 1), L(CI(total_col) + 2)
    colmap_total = {'AI': last_col, 'AJ': total_col, 'AK': agg_col, 'AL': chk_col}
    # geometry
    for c in ('A', 'B'):
        ws.column_dimensions[c].width = M_NOV.column_dimensions[c].width
    for comp in companies:
        ws.column_dimensions[comp['col']].width = M_NOV.column_dimensions['C'].width
    ws.column_dimensions[total_col].width = M_NOV.column_dimensions['AJ'].width
    ws.column_dimensions[agg_col].width = M_NOV.column_dimensions['AK'].width
    ws.column_dimensions[chk_col].width = M_NOV.column_dimensions['AL'].width
    for r in range(1, LAST_ROW + 1):
        if M_NOV.row_dimensions[r].height:
            ws.row_dimensions[r].height = M_NOV.row_dimensions[r].height
    ws.freeze_panes = 'C3'
    ws.sheet_view.zoomScale = 80

    for r in range(1, LAST_ROW + 1):
        # label columns
        for c in ('A', 'B'):
            cstyle(M_NOV[f'{c}{r}'], ws[f'{c}{r}'])
            v = M_NOV[f'{c}{r}'].value
            ws[f'{c}{r}'].value = v
        if r == 1:
            ws['B1'].value = title
        if r == 2:
            ws['B2'].value = month
            ws['A2'].value = 202610 if mi == 0 else \
                f'=VALUE(TEXT(EDATE(DATE(LEFT({MONTHS[mi-1]}!A2,4),RIGHT({MONTHS[mi-1]}!A2,2),1),1),"yyyymm"))'
        # company columns
        for comp in companies:
            col = comp['col']
            cell = ws[f'{col}{r}']
            cstyle(M_NOV[f'C{r}'], cell)
            lock(cell)
            if r == 1:
                cell.value = comp.get('code')
            elif r == 2:
                cell.value = comp['name']
            elif r == 31:
                if mi == 0:
                    cell.value = feed(r, comp, month)
                    if cell.value is None:
                        unlock(cell)
                else:
                    cell.value = f'=IF({MONTHS[mi-1]}!{col}5="","",{MONTHS[mi-1]}!{col}5)'
            elif r in INPUT_ROWS:
                cell.value = feed(r, comp, month)
                if cell.value is None:
                    unlock(cell)
            else:
                cell.value = fixed_company_formula(r, col) or \
                    remap_cols(translate(M_NOV[f'C{r}'].value, f'C{r}', f'{col}{r}'), {'AJ': total_col})
        # TOTAL column
        tc = ws[f'{total_col}{r}']
        cstyle(M_NOV[f'AJ{r}'], tc)
        mv = ftext(M_NOV[f'AJ{r}'].value)
        if r in (46, 47):
            rng = lambda row: f'C{row}:{last_col}{row}'
            tc.value = ArrayFormula(f'{total_col}{r}',
                f'=IF(COUNT({rng(r)})=0,"",IFERROR(SUMPRODUCT(IFERROR({rng(r)}*(({rng(3)})+({rng(4)})),0))'
                f'/SUMPRODUCT(IFERROR(ISNUMBER({rng(r)})*(({rng(3)})+({rng(4)})),0)),""))')
        elif r >= 3 and fixed_company_formula(r, total_col):
            tc.value = fixed_company_formula(r, total_col)
        elif isinstance(mv, str) and mv.startswith('='):
            tc.value = remap_cols(mv, colmap_total)
        else:
            tc.value = mv
        # aggregation label + check columns
        cstyle(M_NOV[f'AK{r}'], ws[f'{agg_col}{r}'])
        ws[f'{agg_col}{r}'].value = M_NOV[f'AK{r}'].value
        cstyle(M_NOV[f'AL{r}'], ws[f'{chk_col}{r}'])
        v = M_NOV[f'AL{r}'].value
        ws[f'{chk_col}{r}'].value = remap_cols(v, colmap_total) if isinstance(v, str) and v.startswith('=') else v
    ws[f'{chk_col}2'].comment = Comment('TRUE = the Headcount = 1.00/0.75/0.50/0.25 split does not add up to '
                                        'Employee Headcount (row 5). Difference shown in row 11.', 'HR Metrics')
    protect(ws)

# ------------------------------------------------------- Quarter tab ------
def build_quarter(ws, title, total_col):
    for c, d in M_Q.column_dimensions.items():
        if d.width:
            ws.column_dimensions[c].width = d.width
    for r in range(1, LAST_ROW + 1):
        if M_Q.row_dimensions[r].height:
            ws.row_dimensions[r].height = M_Q.row_dimensions[r].height
    ws.freeze_panes = 'C3'
    ws.sheet_view.zoomScale = M_Q.sheet_view.zoomScale
    for r in range(1, LAST_ROW + 1):
        for ci in range(1, 20):
            src = M_Q.cell(r, ci); dst = ws.cell(r, ci)
            cstyle(src, dst)
            v = ftext(src.value)
            col = L(ci)
            if r >= 3 and col in Q_MONTH_COL.values():
                month = [m for m, c in Q_MONTH_COL.items() if c == col][0]
                v = f'={month}!{total_col}{r}'                      # FIX: model drifted AJ->AI->AH...
            elif r >= 3 and col in Q_COLS:
                fx = fixed_company_formula(r, col)
                if fx and r not in (68, 69):
                    v = fx
                elif r in (68, 69):
                    den = '5' if r == 68 else '56'
                    v = f'=IF(COUNT({col}58:{col}59)=0,"",IFERROR((N({col}58)+N({col}59))/{col}{den},""))'
                elif r in (39, 40):                                 # YTD rows: latest month reported
                    span = {'F': 'C{r}:E{r}', 'J': 'G{r}:I{r}', 'N': 'K{r}:M{r}', 'R': 'O{r}:Q{r}',
                            'S': 'C{r}:Q{r}'}[col].format(r=r)
                    v = f'=IFERROR(LOOKUP(2,1/ISNUMBER({span}),{span}),"")'
            if r >= 3 and col in Q_COLS:
                dst.number_format = qnf(r)
            if isinstance(src.value, ArrayFormula):
                dst.value = ArrayFormula(f'{col}{r}', v)
            else:
                dst.value = v
    ws['B1'].value = title
    protect(ws)

# ------------------------------------------------------- company tab ------
THIN = Side(style='thin')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_TITLE = Font(name='Calibri', size=14, color='FF000000')
F_RED = Font(name='Calibri', size=14, color='FFC00000')
F_HDR = Font(name='Calibri', size=12, bold=True, color='FFFFFFFF')
F_BODY = Font(name='Calibri', size=12, color='FF31216B')
F_TOT = Font(name='Calibri', size=12, bold=True, color='FF0E2841')
F_NOTE = Font(name='Calibri', size=11, italic=True, color='FFC00000')
FILL_HDR = PatternFill('solid', fgColor='FF595959')
FILL_TITLE = PatternFill('solid', fgColor='FFF2F2F2')
FILL_IN = PatternFill('solid', fgColor='FFBDD7EE')
FILL_TOT = PatternFill('solid', fgColor='FFD9D9D9')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

def put(ws, ref, value=None, font=F_BODY, fill=None, nf=None, align=CENTER, border=BOX, locked=True):
    c = ws[ref]
    c.value = value
    c.font = copy.copy(font)
    if fill: c.fill = copy.copy(fill)
    if nf: c.number_format = nf
    c.alignment = copy.copy(align)
    if border: c.border = copy.copy(border)
    p = copy.copy(c.protection); p.locked = locked; c.protection = p
    return c

def merge(ws, rng, *a, **k):
    first = rng.split(':')[0]
    put(ws, first, *a, **k)
    for row in ws[rng]:
        for c in row:
            if c.coordinate != first:
                c.border = copy.copy(BOX)
    ws.merge_cells(rng)

def build_company_tab(ws, name, col, n_rows):
    """Absence (non-certified / certified / total), Time-to-Hire YTD by month, Time-to-Hire tables."""
    widths = {'A': 5.5, 'B': 16, 'C': 13, 'D': 13, 'E': 14, 'F': 14, 'G': 13, 'H': 5.5, 'I': 16, 'J': 13, 'K': 13,
              'L': 14, 'M': 14, 'N': 13, 'O': 13, 'P': 3, 'Q': 5.5, 'R': 16, 'S': 13, 'T': 13, 'U': 13, 'V': 14,
              'W': 13, 'X': 3, 'Y': 5.5, 'Z': 16, 'AA': 13, 'AB': 13, 'AC': 13, 'AD': 13, 'AE': 13}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    ws.sheet_view.zoomScale = 80
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 24.75
    ws.row_dimensions[2].height = 18.75
    ws.row_dimensions[3].height = 55
    for r in range(4, 17):
        ws.row_dimensions[r].height = 18.75

    # --- absence blocks: (first col, label, days-lost monthly row)
    blocks = [('A', 'NON-CERTIFIED', 43), ('I', 'CERTIFIED', 44), ('Q', 'TOTAL ABSENCES', None)]
    for b, (c0, label, mrow) in enumerate(blocks):
        cols = [L(CI(c0) + i) for i in range(7)]      # No, Month, Employees, WorkDays, DaysLost, TotWD, Rate
        No, Mo, Em, Wd, Dl, Tw, Rt = cols
        merge(ws, f'{No}1:{Em}1', '=A1' if b else name, font=F_TITLE, fill=FILL_TITLE)
        merge(ws, f'{Wd}1:{Tw}2', 'ABSENCE RATE YTD', font=F_TITLE, fill=FILL_TITLE)
        merge(ws, f'{Rt}1:{Rt}2', f'={Rt}16', font=F_TITLE, fill=FILL_TITLE, nf='0.0%')
        merge(ws, f'{No}2:{Em}2', label, font=F_RED, fill=FILL_TITLE)
        for c, h in zip(cols, ['No.', 'Month', 'Total Employees', 'Working Days', 'Total Days Lost',
                               'Total Working Days', 'Monthly Absence Rate']):
            put(ws, f'{c}3', h, font=F_HDR, fill=FILL_HDR)
        for i, m in enumerate(MONTHS):
            r = 4 + i
            put(ws, f'{No}{r}', i + 1)
            put(ws, f'{Mo}{r}', MONTH_DATES[i], nf='mmmm-yy', align=LEFT)
            if b == 0:
                put(ws, f'{Em}{r}', f'=IF({m}!${col}$5="","",{m}!${col}$5)')
                put(ws, f'{Wd}{r}', None, fill=FILL_IN, nf='0', locked=False)
            else:
                put(ws, f'{Em}{r}', f'=C{r}')
                put(ws, f'{Wd}{r}', f'=IF(D{r}="","",D{r})', nf='0')
            if mrow:
                put(ws, f'{Dl}{r}', f'=IF({m}!${col}${mrow}="","",{m}!${col}${mrow})', nf='0.0')
            else:
                put(ws, f'{Dl}{r}', f'=IF(COUNT(E{r},M{r})=0,"",SUM(E{r},M{r}))', nf='0.0')
            put(ws, f'{Tw}{r}', f'=IF(OR({Em}{r}="",{Wd}{r}=""),"",{Em}{r}*{Wd}{r})', nf='#,##0')
            put(ws, f'{Rt}{r}', f'=IF(OR({Dl}{r}="",N({Tw}{r})=0),"",{Dl}{r}/{Tw}{r})', nf='0.00%')
        put(ws, f'{Wd}16', 'Total YTD', font=F_TOT, fill=FILL_TOT, align=Alignment(horizontal='right'))
        put(ws, f'{Dl}16', f'=IF(COUNT({Dl}4:{Dl}15)=0,"",SUM({Dl}4:{Dl}15))', font=F_TOT, fill=FILL_TOT, nf='0.0')
        put(ws, f'{Tw}16', f'=SUMPRODUCT(--ISNUMBER({Dl}4:{Dl}15),{Tw}4:{Tw}15)', font=F_TOT, fill=FILL_TOT, nf='#,##0')
        put(ws, f'{Rt}16', f'=IF(COUNT({Dl}4:{Dl}15)=0,"",IFERROR({Dl}16/{Tw}16,""))', font=F_TOT, fill=FILL_TOT, nf='0.00%')
        for c in (No, Mo, Em):
            put(ws, f'{c}16', None, fill=FILL_TOT)

    # --- Time-to-Hire tables (rows 18..)
    top, first = 18, 21
    last = first + n_rows - 1
    tot = last + 1
    ws.row_dimensions[18].height = 24.75
    ws.row_dimensions[19].height = 21.75
    ws.row_dimensions[20].height = 60
    tth = [('A', 'PERMANENT EE'), ('H', 'TEMPORARY EE')]
    for k, (c0, label) in enumerate(tth):
        No, Jt, Po, Od, Sd, Dy, Pf = [L(CI(c0) + i) for i in range(7)]
        merge(ws, f'{No}18:{Po}18', '=A1', font=F_TITLE, fill=FILL_TITLE)
        merge(ws, f'{Od}18:{Sd}19', 'TIME TO HIRE YTD (Days)', font=F_TITLE, fill=FILL_TITLE)
        merge(ws, f'{Dy}18:{Pf}19', f'={"AB" if k == 0 else "AD"}16', font=F_TITLE, fill=FILL_TITLE, nf='0.0')
        merge(ws, f'{No}19:{Po}19', label, font=F_RED, fill=FILL_TITLE)
        for c, h in zip([No, Jt, Po, Od, Sd, Dy, Pf],
                        ['No.', 'Job Title', 'No. of Posts', 'Job Opening Date (mm/dd/yyyy)',
                         'Contract Signed Date (mm/dd/yyyy)', 'Time to Hire by Position (Days)', 'Posts Filled (auto)']):
            put(ws, f'{c}20', h, font=F_HDR, fill=FILL_HDR)
        for r in range(first, last + 1):
            put(ws, f'{No}{r}', r - first + 1)
            put(ws, f'{Jt}{r}', None, fill=FILL_IN, align=LEFT, locked=False)
            put(ws, f'{Po}{r}', None, fill=FILL_IN, nf='0', locked=False)
            put(ws, f'{Od}{r}', None, fill=FILL_IN, nf='mm/dd/yyyy', locked=False)
            put(ws, f'{Sd}{r}', None, fill=FILL_IN, nf='mm/dd/yyyy', locked=False)
            put(ws, f'{Dy}{r}', f'=IF(OR({Sd}{r}="",{Od}{r}=""),"",MAX({Sd}{r}-{Od}{r},0))', nf='0')
            put(ws, f'{Pf}{r}', f'=IF({Dy}{r}="","",IF(N({Po}{r})=0,1,{Po}{r}))', nf='0')
        put(ws, f'{Jt}{tot}', 'Total', font=F_TOT, fill=FILL_TOT, align=Alignment(horizontal='right'))
        put(ws, f'{Po}{tot}', f'=SUM({Po}{first}:{Po}{last})', font=F_TOT, fill=FILL_TOT, nf='0')
        put(ws, f'{Dy}{tot}', f'=IF(N({Pf}{tot})=0,"",SUMPRODUCT({Dy}{first}:{Dy}{last},{Pf}{first}:{Pf}{last})/{Pf}{tot})',
            font=F_TOT, fill=FILL_TOT, nf='0.0')
        put(ws, f'{Pf}{tot}', f'=SUM({Pf}{first}:{Pf}{last})', font=F_TOT, fill=FILL_TOT, nf='0')
        for c in (No, Od, Sd):
            put(ws, f'{c}{tot}', None, fill=FILL_TOT)
        dv_d = DataValidation(type='date', operator='between', formula1='DATE(2020,1,1)', formula2='DATE(2030,12,31)',
                              allow_blank=True, showErrorMessage=True, errorTitle='Date required',
                              error='Enter a date (mm/dd/yyyy).')
        dv_d.add(f'{Od}{first}:{Sd}{last}')
        dv_n = DataValidation(type='whole', operator='greaterThanOrEqual', formula1='1', allow_blank=True,
                              showErrorMessage=True, error='No. of Posts must be a whole number >= 1.')
        dv_n.add(f'{Po}{first}:{Po}{last}')
        ws.add_data_validation(dv_d); ws.add_data_validation(dv_n)
    ws.cell(tot + 1, 2).value = ("Insert rows above the Total row as necessary and copy the formulas of the "
                                 "'Time to Hire' and 'Posts Filled' columns into the new rows. "
                                 "Blank 'No. of Posts' on a filled position counts as 1 post.")
    ws.cell(tot + 1, 2).font = F_NOTE
    dv_w = DataValidation(type='whole', operator='between', formula1='0', formula2='31', allow_blank=True,
                          showErrorMessage=True, error='Working days must be between 0 and 31.')
    dv_w.add('D4:D15'); ws.add_data_validation(dv_w)

    # --- Time-to-Hire YTD by month (feeds monthly tabs rows 39/40)
    E, G, Lc, N_ = (f'$E${first}:$E${last}', f'$G${first}:$G${last}', f'$L${first}:$L${last}', f'$N${first}:$N${last}')
    F_, M_ = f'$F${first}:$F${last}', f'$M${first}:$M${last}'
    merge(ws, 'Y1:AA1', '=A1', font=F_TITLE, fill=FILL_TITLE)
    merge(ws, 'Y2:AA2', 'TIME TO HIRE BY MONTH', font=F_RED, fill=FILL_TITLE)
    merge(ws, 'AB1:AD2', 'TIME TO HIRE YTD (Days)', font=F_TITLE, fill=FILL_TITLE)
    merge(ws, 'AE1:AE2', '=AE16', font=F_TITLE, fill=FILL_TITLE, nf='0.0')
    for c, h in zip(['Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE'],
                    ['No.', 'Month', 'Permanent Posts Filled YTD', 'Permanent Time to Hire YTD (Days)',
                     'Temporary Posts Filled YTD', 'Temporary Time to Hire YTD (Days)', 'Total Time to Hire YTD (Days)']):
        put(ws, f'{c}3', h, font=F_HDR, fill=FILL_HDR)
    def ytd(r, end, sd, pf, dy=None):
        win = f'({sd}>=$Z$4)*({sd}<=EOMONTH({end},0))'
        return f'SUMPRODUCT({win},{dy},{pf})' if dy else f'SUMPRODUCT({win},{pf})'
    for i in range(12):
        r = 4 + i
        put(ws, f'Y{r}', i + 1)
        put(ws, f'Z{r}', MONTH_DATES[i], nf='mmmm-yy', align=LEFT)
        put(ws, f'AA{r}', f'=IF($C{r}="","",{ytd(r, f"$Z{r}", E, G)})', nf='0')
        put(ws, f'AB{r}', f'=IF(N(AA{r})=0,"",{ytd(r, f"$Z{r}", E, G, F_)}/AA{r})', nf='0.0')
        put(ws, f'AC{r}', f'=IF($C{r}="","",{ytd(r, f"$Z{r}", Lc, N_)})', nf='0')
        put(ws, f'AD{r}', f'=IF(N(AC{r})=0,"",{ytd(r, f"$Z{r}", Lc, N_, M_)}/AC{r})', nf='0.0')
        put(ws, f'AE{r}', f'=IF(N(AA{r})+N(AC{r})=0,"",(N(AB{r})*N(AA{r})+N(AD{r})*N(AC{r}))/(N(AA{r})+N(AC{r})))', nf='0.0')
    put(ws, 'Y16', None, fill=FILL_TOT)
    put(ws, 'Z16', 'FY2027 YTD', font=F_TOT, fill=FILL_TOT, align=Alignment(horizontal='right'))
    put(ws, 'AA16', f'={ytd(16, "$Z$15", E, G)}', font=F_TOT, fill=FILL_TOT, nf='0')
    put(ws, 'AB16', f'=IF(N(AA16)=0,"",{ytd(16, "$Z$15", E, G, F_)}/AA16)', font=F_TOT, fill=FILL_TOT, nf='0.0')
    put(ws, 'AC16', f'={ytd(16, "$Z$15", Lc, N_)}', font=F_TOT, fill=FILL_TOT, nf='0')
    put(ws, 'AD16', f'=IF(N(AC16)=0,"",{ytd(16, "$Z$15", Lc, N_, M_)}/AC16)', font=F_TOT, fill=FILL_TOT, nf='0.0')
    put(ws, 'AE16', '=IF(N(AA16)+N(AC16)=0,"",(N(AB16)*N(AA16)+N(AD16)*N(AC16))/(N(AA16)+N(AC16)))',
        font=F_TOT, fill=FILL_TOT, nf='0.0')
    ws['Y17'].value = 'Only contracts signed within FY2027 (Oct-2026 to the month shown) are counted; weighted by posts filled.'
    ws['Y17'].font = F_NOTE
    protect(ws, rows_editable=True)

# --------------------------------------------------- entity workbook ------
def build_entity_workbook(title, companies, path, fy26_file_for):
    """Cluster / division workbook: Quarter + 12 monthly tabs + company tabs."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    total_col = L(CI(companies[-1]['col']) + 1)
    qs = wb.create_sheet('Quarter')
    month_ws = [wb.create_sheet(m) for m in MONTHS]

    def feed(r, comp, month):
        if r in TAB_ROWS:
            mi = MONTHS.index(month)
            return blank_if(f"{q(comp['tab'])}!${TAB_ROWS[r]}${4 + mi}")
        return None

    for mi, ws in enumerate(month_ws):
        build_month(ws, mi, title, companies, total_col, feed)
    build_quarter(qs, title, total_col)
    for comp in companies:
        ws = wb.create_sheet(comp['tab'])
        build_company_tab(ws, comp['name'], comp['col'], tth_rows(*fy26_file_for(comp)))
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return dict(path=path, total_col=total_col, companies=companies)

# ------------------------------------------------------------- build ------
def main():
    built = {}
    model_names = {L(c): (M_NOV.cell(1, c).value, M_NOV.cell(2, c).value) for c in range(3, 36)}

    # 1) GCG clusters
    for key, title, fname in CLUSTERS:
        cols = [c for c, v in GCG_MAP.items() if v[0] == key]
        comps = []
        for i, mc in enumerate(cols):
            code, name = model_names[mc]
            comps.append(dict(col=L(3 + i), code=code, name=name, tab=GCG_MAP[mc][1], fy26=GCG_MAP[mc][2], model_col=mc))
        path = os.path.join(OUT, GCG_DIR, fname)
        built[key] = build_entity_workbook(title, comps, path, lambda c, k=key: (FY26[k], c['fy26']))
        print('built', path)

    # 2) 0 GCG (33 Companies) FY2027 - model layout, monthly inputs linked to the clusters
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    gcg_path = os.path.join(OUT, GCG_DIR, GCG_FILE)
    link_idx = {}
    for key, _, fname in CLUSTERS:
        link_idx[key] = add_ext_link(wb, rel_url(os.path.dirname(gcg_path), built[key]['path']),
                                     ['Quarter'] + MONTHS + [c['tab'] for c in built[key]['companies']])
    comps = [dict(col=mc, code=model_names[mc][0], name=model_names[mc][1]) for mc in GCG_MAP]
    where = {c['model_col']: (k, c['col']) for k in built if k in dict((x[0], 1) for x in CLUSTERS)
             for c in built[k]['companies']}

    def feed_gcg(r, comp, month):
        k, ccol = where[comp['col']]
        return blank_if(f"[{link_idx[k]}]{month}!{ccol}{r}")
    qs = wb.create_sheet('Quarter')
    for mi, m in enumerate(MONTHS):
        build_month(wb.create_sheet(m), mi, 'GCG', comps, 'AJ', feed_gcg)
    build_quarter(qs, 'GCG', 'AJ')
    wb.calculation.fullCalcOnLoad = True
    wb.save(gcg_path)
    built['GCG'] = dict(path=gcg_path)
    print('built', gcg_path)

    # 3) Divisions
    for title, folder, fname, fy26, comp_list in DIVISIONS:
        comps = [dict(col=L(3 + i), code=None, name=n, tab=t, fy26=t26) for i, (n, t, t26) in enumerate(comp_list)]
        path = os.path.join(OUT, folder, fname)
        built[title] = build_entity_workbook(title, comps, path, lambda c, f=fy26: (f, c['fy26']))
        print('built', path)

    # 4) GEL HR Metrics FY2027
    build_gel(built)

def build_gel(built):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = 'GEL Divisions'
    gel_path = os.path.join(OUT, GEL_DIR, GEL_FILE)
    divs = [(t, built[t]['path']) for t, *_ in DIVISIONS] + [('GCG', built['GCG']['path'])]
    idx = {}
    for t, p in divs:
        idx[t] = add_ext_link(wb, rel_url(os.path.dirname(gel_path), p), ['Quarter'] + MONTHS)
    ncol = 5
    div_cols = {t: [L(3 + d * ncol + i) for i in range(ncol)] for d, (t, _) in enumerate(divs)}
    gel_cols = [L(3 + len(divs) * ncol + i) for i in range(ncol)]
    agg_col = L(CI(gel_cols[-1]) + 1)
    titles = {'AUTOMOTIVE': 'AUTO', 'GCG': 'GCG (33 COMPANIES)'}

    ws.column_dimensions['A'].width = M_Q.column_dimensions['A'].width
    ws.column_dimensions['B'].width = M_Q.column_dimensions['B'].width
    for c in [c for cs in div_cols.values() for c in cs] + gel_cols:
        ws.column_dimensions[c].width = 13.7
    ws.column_dimensions[agg_col].width = 16
    ws.freeze_panes = 'C3'
    ws.sheet_view.zoomScale = 80
    for r in range(1, LAST_ROW + 1):
        if M_Q.row_dimensions[r].height:
            ws.row_dimensions[r].height = M_Q.row_dimensions[r].height

    # header rows
    for r in (1, 2):
        for c in ('A', 'B'):
            cstyle(M_Q[f'{c}{r}'], ws[f'{c}{r}'])
    ws['B1'].value = 'GEL DIVISIONS'; ws['A2'].value = 'FY2027'; ws['B2'].value = 'Quarter'
    blocks = [(titles.get(t, t), cs) for t, cs in div_cols.items()] + [('GEL', gel_cols)]
    for name, cs in blocks:
        for i, c in enumerate(cs):
            cstyle(M_Q['B1'], ws[f'{c}1'])
            cstyle(M_Q['S2'] if i == 4 else M_Q['F2'], ws[f'{c}2'])
            ws[f'{c}2'].value = Q_LABELS[i]
        ws[f'{cs[0]}1'].value = name
        ws.merge_cells(f'{cs[0]}1:{cs[-1]}1')
    cstyle(M_NOV['AK2'], ws[f'{agg_col}2']); ws[f'{agg_col}2'].value = 'Consolidation'
    cstyle(M_NOV['AK1'], ws[f'{agg_col}1'])

    for r in range(FIRST_ROW, LAST_ROW + 1):
        cstyle(M_Q[f'A{r}'], ws[f'A{r}']); cstyle(M_Q[f'B{r}'], ws[f'B{r}'])
        ws[f'B{r}'].value = M_Q[f'B{r}'].value
        agg = AGG[r]
        for t, cs in div_cols.items():
            for i, c in enumerate(cs):
                cell = ws[f'{c}{r}']
                cstyle(M_Q[f'{"S" if i == 4 else "F"}{r}'], cell)
                cell.number_format = qnf(r)
                if r == 49:
                    cell.value = f'=IFERROR({c}5/{gel_cols[i]}5,"")'
                else:
                    cell.value = blank_if(f'[{idx[t]}]Quarter!{Q_COLS[i]}{r}')
        for i, g in enumerate(gel_cols):
            cell = ws[f'{g}{r}']
            cstyle(M_NOV[f'AJ{r}'], cell)
            cell.number_format = qnf(r)
            parts = [div_cols[t][i] + str(r) for t in div_cols]
            w = [div_cols[t][i] + '5' for t in div_cols]
            qf = M_Q[f'F{r}'].value
            qf = ftext(qf)
            refs_only_f = isinstance(qf, str) and all(m.group(2) == 'F' for m in CELL_RE.finditer(qf))
            if r in (46, 47, 65):
                num = '+'.join(f'IFERROR({p}*{x},0)' for p, x in zip(parts, w))
                den = '+'.join(f'IFERROR(ISNUMBER({p})*{x},0)' for p, x in zip(parts, w))
                cell.value = f'=IF(COUNT({",".join(parts)})=0,"",IFERROR(({num})/({den}),""))'
            elif agg == 'Avg':
                cell.value = f'=IFERROR(AVERAGE({",".join(parts)}),"")'
            elif r == 49 or agg == '∑' or r == 14:
                cell.value = f'=IF(COUNT({",".join(parts)})=0,"",SUM({",".join(parts)}))'
            elif refs_only_f:
                fx = fixed_company_formula(r, 'F') if r not in (68, 69) else \
                    f'=IF(COUNT(F58:F59)=0,"",IFERROR((N(F58)+N(F59))/F{"5" if r == 68 else "56"},""))'
                cell.value = remap_cols(fx or qf, {'F': g})
            else:
                raise ValueError(f'GEL: no consolidation rule for row {r}')
        a = ws[f'{agg_col}{r}']
        cstyle(M_NOV[f'AK{r}'], a)
        a.value = {'∑': 'Sum of divisions', 'Avg': 'Average of divisions', 'Weight Factor': 'Division HC / GEL HC',
                   'Weighted': 'HC-weighted'}.get(agg, 'Recalculated')
        if r == 14: a.value = 'Sum of divisions'
        if r == 65: a.value = 'HC-weighted'
    protect(ws)
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(gel_path), exist_ok=True)
    wb.save(gel_path)
    print('built', gel_path)

if __name__ == '__main__':
    main()
