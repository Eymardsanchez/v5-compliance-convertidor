"""
Extractor para estados de cuenta HSBC (Cuenta de Cheques).
El PDF usa fuentes Type3 con encoding completamente personalizado (256 entradas).
Se usa la tabla DIFFERENCES del font para construir un mapa de decodificación.
Extracción basada en palabras con bounding-boxes para clasificar columnas exactas:
  Retiro/Cargo  → x0 ∈ [345, 420)
  Depósito/Abono → x0 ∈ [420, 510)
  Saldo          → x0 ≥ 510
"""

import re
from typing import Optional, Callable

import pdfplumber

from bancos.base import BancoExtractor, Movimiento, InfoCuenta

# ── Mapa de decodificación de la fuente Type3 de HSBC ──────────────────────────
# Differences[código] = 'EX{N}000'  → el glifo dibuja chr(N)
_DIFFERENCES = {
    0:'EX000000',1:'EX032000',2:'EX032000',3:'EX032000',4:'EX032000',
    5:'EX032000',6:'EX032000',7:'EX032000',8:'EX032000',9:'EX032000',
    10:'EX133000',11:'EX032000',12:'EX032000',13:'EX032000',14:'EX032000',
    15:'EX032000',16:'EX032000',17:'EX032000',18:'EX032000',19:'EX032000',
    20:'EX032000',21:'EX032000',22:'EX032000',23:'EX032000',24:'EX032000',
    25:'EX032000',26:'EX032000',27:'EX032000',28:'EX032000',29:'EX032000',
    30:'EX030000',31:'EX031000',32:'EX128000',33:'EX129000',34:'EX127000',
    35:'EX131000',36:'EX132000',37:'EX032000',38:'EX134000',39:'EX135000',
    40:'EX136000',41:'EX137000',42:'EX138000',43:'EX139000',44:'EX140000',
    45:'EX184000',46:'EX142000',47:'EX143000',48:'EX144000',49:'EX145000',
    50:'EX146000',51:'EX147000',52:'EX148000',53:'EX183000',54:'EX150000',
    55:'EX151000',56:'EX152000',57:'EX153000',58:'EX154000',59:'EX155000',
    60:'EX156000',61:'EX157000',62:'EX158000',63:'EX159000',64:'EX032000',
    65:'EX141000',66:'EX226000',67:'EX228000',68:'EX224000',69:'EX225000',
    70:'EX227000',71:'EX229000',72:'EX231000',73:'EX241000',74:'EX162000',
    75:'EX046000',76:'EX060000',77:'EX040000',78:'EX043000',79:'EX124000',
    80:'EX038000',81:'EX233000',82:'EX234000',83:'EX235000',84:'EX232000',
    85:'EX237000',86:'EX238000',87:'EX239000',88:'EX236000',89:'EX223000',
    90:'EX033000',91:'EX036000',92:'EX042000',93:'EX041000',94:'EX059000',
    95:'EX172000',96:'EX045000',97:'EX047000',98:'EX194000',99:'EX196000',
    100:'EX192000',101:'EX193000',102:'EX195000',103:'EX197000',104:'EX199000',
    105:'EX209000',106:'EX166000',107:'EX044000',108:'EX037000',109:'EX095000',
    110:'EX062000',111:'EX063000',112:'EX160000',113:'EX201000',114:'EX202000',
    115:'EX203000',116:'EX200000',117:'EX205000',118:'EX206000',119:'EX207000',
    120:'EX204000',121:'EX096000',122:'EX058000',123:'EX035000',124:'EX064000',
    125:'EX039000',126:'EX061000',127:'EX034000',128:'EX216000',129:'EX097000',
    130:'EX098000',131:'EX099000',132:'EX100000',133:'EX101000',134:'EX102000',
    135:'EX103000',136:'EX104000',137:'EX105000',138:'EX171000',139:'EX187000',
    140:'EX240000',141:'EX253000',142:'EX222000',143:'EX177000',144:'EX176000',
    145:'EX106000',146:'EX107000',147:'EX108000',148:'EX109000',149:'EX110000',
    150:'EX111000',151:'EX112000',152:'EX113000',153:'EX114000',154:'EX170000',
    155:'EX186000',156:'EX230000',157:'EX130000',158:'EX198000',159:'EX164000',
    160:'EX181000',161:'EX126000',162:'EX115000',163:'EX116000',164:'EX117000',
    165:'EX118000',166:'EX119000',167:'EX120000',168:'EX121000',169:'EX122000',
    170:'EX161000',171:'EX191000',172:'EX208000',173:'EX221000',174:'EX254000',
    175:'EX174000',176:'EX094000',177:'EX163000',178:'EX165000',179:'EX149000',
    180:'EX169000',181:'EX167000',182:'EX182000',183:'EX188000',184:'EX189000',
    185:'EX190000',186:'EX091000',187:'EX093000',188:'EX173000',189:'EX168000',
    190:'EX180000',191:'EX215000',192:'EX123000',193:'EX065000',194:'EX066000',
    195:'EX067000',196:'EX068000',197:'EX069000',198:'EX070000',199:'EX071000',
    200:'EX072000',201:'EX073000',202:'EX175000',203:'EX244000',204:'EX246000',
    205:'EX242000',206:'EX243000',207:'EX245000',208:'EX125000',209:'EX074000',
    210:'EX075000',211:'EX076000',212:'EX077000',213:'EX078000',214:'EX079000',
    215:'EX080000',216:'EX081000',217:'EX082000',218:'EX185000',219:'EX251000',
    220:'EX252000',221:'EX249000',222:'EX250000',223:'EX255000',224:'EX092000',
    225:'EX247000',226:'EX083000',227:'EX084000',228:'EX085000',229:'EX086000',
    230:'EX087000',231:'EX088000',232:'EX089000',233:'EX090000',234:'EX178000',
    235:'EX212000',236:'EX214000',237:'EX210000',238:'EX211000',239:'EX213000',
    240:'EX048000',241:'EX049000',242:'EX050000',243:'EX051000',244:'EX052000',
    245:'EX053000',246:'EX054000',247:'EX055000',248:'EX056000',249:'EX057000',
    250:'EX179000',251:'EX219000',252:'EX220000',253:'EX217000',254:'EX218000',
    255:'EX248000',
}

# Construir tabla de decodificación: código → carácter real
_DECODE: dict[int, str] = {}
for _code, _glyph in _DIFFERENCES.items():
    _m = re.match(r'EX(\d+)000', _glyph)
    if _m:
        _n = int(_m.group(1))
        _DECODE[_code] = chr(_n) if _n > 0 else ' '
    else:
        _DECODE[_code] = ' '

# Mapa inverso: chr(código) → carácter real  (para chars que pdfminer pasa como Latin-1)
_PDFMINER_TO_ACTUAL: dict[str, str] = {
    chr(c): _DECODE.get(c, chr(c)) for c in range(256)
}

# Correcciones de dígitos para chars de Unicode alto (pdfminer glyph-name lookup)
_DIGIT_FIX = str.maketrans({'W': '1', 'ı': '5', 'ł': '8'})

# Tabla de corrección de caracteres de descripción:
# pdfminer resuelve nombres de glifos estándar (circumflex→ˆ, registered→®, etc.)
# usando la Adobe Glyph List, en lugar de pasar (cid:N). Estos mapeos son específicos
# de la fuente Type3 de HSBC y deben revertirse al carácter real.
_HSBC_CHAR_FIX = str.maketrans({
    'ˆ': 'C',   # U+02C6 circumflex modifier → C  (EX067=C)
    '®': 'E',   # U+00AE registered sign    → E  (EX069=E)
    '¡': 'T',   # U+00A1 inv. exclamation   → T  (EX084=T)
    '˜': 'D',   # U+02DC small tilde        → D  (EX068=D)
    '˙': 'G',   # U+02D9 dot above          → G  (EX071=G)
    '˘': 'F',   # U+02D8 breve              → F  (EX070=F)
    '©': 'B',   # U+00A9 copyright sign     → B  (EX066=B)
    'Ł': 'Y',   # U+0141 L with stroke      → Y  (EX089=Y)
    'y': 'H',   # ASCII y in HSBC encoding  → H  (EX072=H)
    '⁄': 'U',   # U+2044 fraction slash     → U  (EX085=U)
    '˛': 'O',   # U+02DB ogonek             → O  (EX079=O)
    '-': 'A',  # hyphen-minus          → A  (EX065=A)
    'ł': 'S',   # U+0142 l with stroke      → S  (EX083=S) — dígito 8 sólo en numérico
    'ı': 'I',   # U+0131 dotless i          → I  (EX073=I) — dígito 5 sólo en numérico
    'W': 'N',   # ASCII W en HSBC encoding  → N  (EX078=N) — dígito 1 sólo en numérico
    'Ð': 'R',   # U+00D0 Eth               → R
    'ð': 'r',   # U+00F0 eth               → r (minúscula)
    'Ø': 'P',   # U+00D8 O-slash           → P
    'Þ': 'M',   # U+00DE Thorn             → M
    'þ': 'L',   # U+00FE thorn lc          → L
    'Æ': 'V',   # U+00C6 AE               → V
    'æ': 'v',   # U+00E6 ae               → v
})

_PERIOD_RE  = re.compile(r'(\d{2}/\d{2}/\d{4})\s+al\s+(\d{2}/\d{2}/\d{4})')
_AMOUNT_RE  = re.compile(r'^[\d,]+\.\d{2}$')
_DAY_RE     = re.compile(r'^\d{1,2}$')

# Umbrales de columna (página de 612 pt ancho)
_X_DAY_MAX  = 58    # la palabra del día termina antes de x=58
_X_DESC_MAX = 300   # descripción: x0 < 300
_X_REF_MAX  = 345   # referencia: x0 ∈ [300, 345)
_X_RET_MAX  = 420   # Retiro/Cargo: x0 ∈ [345, 420)
_X_DEP_MAX  = 510   # Depósito/Abono: x0 ∈ [420, 510)
#                     Saldo: x0 >= 510


def _decode_word(raw: str) -> str:
    """Decodifica un token de texto de la fuente Type3 de HSBC."""
    result: list[str] = []
    i = 0
    while i < len(raw):
        if raw[i:i+5] == '(cid:':
            end = raw.find(')', i)
            if end != -1:
                n = int(raw[i+5:end])
                result.append(_DECODE.get(n, '?'))
                i = end + 1
                continue
        c = raw[i]
        result.append(_PDFMINER_TO_ACTUAL.get(c, c))
        i += 1
    return ''.join(result)


def _fix_num(s: str) -> str:
    """Sustituye W→1, ı→5, ł→8 en contextos numéricos."""
    return s.translate(_DIGIT_FIX)


def _parse_amount(text: str) -> Optional[float]:
    t = _fix_num(text.replace(',', '').strip())
    try:
        return float(t)
    except ValueError:
        return None


def _group_rows(words: list[dict]) -> dict[int, list[dict]]:
    rows: dict[int, list[dict]] = {}
    for w in words:
        y = round(w['top'] / 2) * 2
        rows.setdefault(y, []).append(w)
    return rows


class HSBCExtractor(BancoExtractor):
    """Extractor para estados de cuenta HSBC Cuenta de Cheques."""

    NOMBRE    = "HSBC"
    COLOR_HEX = "DB0011"   # Rojo corporativo HSBC

    @classmethod
    def detectar(cls, texto: str) -> bool:
        # Los PDF de HSBC usan fuentes Type3 → pdfminer produce muchos (cid:N)
        return '(cid:' in texto or '(CID:' in texto

    # ── extracción principal ──────────────────────────────────────────────────

    def extraer(
        self,
        pdf_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> tuple[list[Movimiento], InfoCuenta]:

        info = InfoCuenta(banco="HSBC")
        movimientos: list[Movimiento] = []
        mes = "01"
        anio = "2018"

        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)

            # ── Paso 1: extraer metadatos de la página 1 ─────────────────────
            raw0 = pdf.pages[0].extract_text(x_tolerance=2, y_tolerance=3) or ""
            decoded0 = _decode_word(raw0)
            fixed0   = _fix_num(decoded0)

            m = _PERIOD_RE.search(fixed0)
            if m:
                d_ini = m.group(1)   # "01/01/2018"
                d_fin = m.group(2)
                parts = d_ini.split('/')
                mes   = parts[1]
                anio  = parts[2]
                info.periodo = f"{d_ini} al {d_fin}"

            # Saldo Inicial  (el símbolo $ puede decodificarse distinto; usamos \D*)
            si = re.search(r'Saldo\s+Inicial\D{0,10}([\d,]+\.\d{2})', fixed0)
            if si:
                info.saldo_inicial = float(si.group(1).replace(',', ''))

            # Saldo Final
            sf = re.search(r'Saldo\s+.inal\D{0,10}([\d,]+\.\d{2})', fixed0)
            if sf:
                info.saldo_final = float(sf.group(1).replace(',', ''))

            # Número de cuenta y CLABE (búsqueda por patrón numérico)
            acc = re.search(r'\b(\d{10})\b', fixed0)
            if acc:
                info.numero_cuenta = acc.group(1)
            clabe = re.search(r'\b(\d{18})\b', fixed0)
            if clabe:
                info.clabe = clabe.group(1)

            # ── Paso 2: extraer movimientos página a página ───────────────────
            for pnum, page in enumerate(pdf.pages):
                if progress_callback:
                    progress_callback(pnum + 1, total)

                words = page.extract_words(x_tolerance=2, y_tolerance=3)
                rows  = _group_rows(words)

                for y in sorted(rows.keys()):
                    ws = sorted(rows[y], key=lambda w: w['x0'])
                    mov = self._parse_row(ws, mes, anio)
                    if mov is not None:
                        movimientos.append(mov)

        if movimientos:
            info.total_depositos = sum(m.deposito or 0 for m in movimientos)
            info.total_retiros   = sum(m.retiro   or 0 for m in movimientos)
            if info.saldo_final is None:
                info.saldo_final = movimientos[-1].saldo

        return movimientos, info

    # ── parsing de una fila de palabras ──────────────────────────────────────

    def _parse_row(
        self,
        ws: list[dict],
        mes: str,
        anio: str,
    ) -> Optional[Movimiento]:
        """
        Retorna un Movimiento si la fila es una fila de transacción válida:
        primer token = 1-2 dígitos (día) en la columna de día.
        """
        if not ws:
            return None

        first = ws[0]
        first_text = _fix_num(_decode_word(first['text'])).strip()

        # El día debe estar en la columna izquierda y ser 1-2 dígitos
        if first['x0'] > _X_DAY_MAX:
            return None
        if not _DAY_RE.match(first_text):
            return None

        day   = first_text.zfill(2)
        fecha = f"{day}/{mes}/{anio}"

        desc_parts: list[str] = []
        retiro:    Optional[float] = None
        deposito:  Optional[float] = None
        saldo:     Optional[float] = None

        for w in ws[1:]:
            x0   = w['x0']
            text = _fix_num(_decode_word(w['text'])).strip()

            if not text or text == '$':
                continue

            if x0 >= _X_DEP_MAX:
                # Columna Saldo
                v = _parse_amount(text)
                if v is not None:
                    saldo = v

            elif x0 >= _X_RET_MAX:
                # Columna Depósito/Abono
                v = _parse_amount(text)
                if v is not None:
                    deposito = v

            elif x0 >= _X_REF_MAX:
                # Columna Retiro/Cargo
                v = _parse_amount(text)
                if v is not None:
                    retiro = v
                # else: podría ser el signo $ previo, se ignora

            elif x0 > _X_DAY_MAX:
                # Descripción o referencia — aplicar corrección de caracteres HSBC
                fixed_text = text.translate(_HSBC_CHAR_FIX)
                if x0 < _X_DESC_MAX:
                    desc_parts.append(fixed_text)
                # else: referencia numérica → incluir como parte de desc si no es serial
                else:
                    # números puros cortos son seriales → ignorar
                    if not re.match(r'^\d{1,8}$', fixed_text):
                        desc_parts.append(fixed_text)

        # Necesitamos al menos saldo para considerar válida la fila
        if saldo is None:
            return None

        desc = ' '.join(desc_parts).strip()
        desc = re.sub(r'\s+', ' ', desc)

        return Movimiento(
            fecha=fecha,
            descripcion=desc,
            deposito=deposito,
            retiro=retiro,
            saldo=saldo,
        )
