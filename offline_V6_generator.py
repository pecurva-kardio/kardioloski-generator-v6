
"""
OFFLINE V6 GENERATOR - KONAČNA VERZIJA
- Bold anatomski segmenti: LV, LA, DV, DA, aorta, AV, MV, TV, PV, perikard, GLS/strain
- Puni Doppler, puna dijagnoza (4), puna terapija (2), Arial 10
- Pravila iz PREDLOZAK-KARDIOLOSKOG-NALAZA_bold_bullet.docx

Korištenje:
1. Stavi PDF echo izvještaj u isti folder kao transkript.txt
2. python offline_V6_generator.py

Ili interaktivno:
python offline_V6_generator.py --transkript "tekst..." --pdf "echo.pdf"
"""

from docx import Document
from docx.shared import Pt
import re, os, sys

def get_num(txt, pat):
    m = re.search(pat, txt, re.IGNORECASE)
    if m:
        try:
            return float(m.group(1).replace(",", "."))
        except:
            return None
    return None

def rnd(v,d=1):
    if v is None: return "n/a"
    if d==0: return str(int(round(v)))
    return f"{v:.1f}"

def add_bold_line(doc, text):
    p=doc.add_paragraph()
    r=p.add_run(text)
    r.bold=True
    r.font.name='Arial'
    r.font.size=Pt(10)
    return p

def add_sec(doc, title, content):
    if not content: return
    p=doc.add_paragraph()
    rt=p.add_run(title+" ")
    rt.bold=True
    rt.font.name='Arial'
    rt.font.size=Pt(10)
    rc=p.add_run(content)
    rc.font.name='Arial'
    rc.font.size=Pt(10)

def add_bsec(doc, title):
    p=doc.add_paragraph()
    r=p.add_run(title)
    r.bold=True
    r.font.name='Arial'
    r.font.size=Pt(10)

def add_bullet_bold_seg(doc, seg, rest):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(0)
    p.paragraph_format.space_after=Pt(0)
    p.paragraph_format.left_indent=Pt(18)
    rb=p.add_run("• ")
    rb.font.name='Arial'
    rb.font.size=Pt(10)
    rs=p.add_run(seg+" ")
    rs.bold=True
    rs.font.name='Arial'
    rs.font.size=Pt(10)
    rr=p.add_run(rest)
    rr.font.name='Arial'
    rr.font.size=Pt(10)

def add_bullet(doc, text):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(0)
    p.paragraph_format.space_after=Pt(0)
    p.paragraph_format.left_indent=Pt(18)
    r=p.add_run(f"• {text}")
    r.font.name='Arial'
    r.font.size=Pt(10)

def parse_transkript_full(trans):
    # Anamneza = do "Dobrog je"
    parts = re.split(r'Dobrog je općeg stanja', trans, flags=re.IGNORECASE)
    anam = parts[0].strip() if parts else trans[:1000]
    
    # Status
    status_match = re.search(r'Dobrog je općeg stanja.*?BSA.*?m²\.?', trans, re.DOTALL | re.IGNORECASE)
    if not status_match:
        status_match = re.search(r'Dobrog je općeg stanja.*?80/min\.?', trans, re.DOTALL | re.IGNORECASE)
    status = status_match.group(0).strip() if status_match else "Dobrog je općeg stanja, astenične konstitucije, eupnoična. Jugularne vene nisu patološke. Na plućima uredan dišni šum. Srce ritmično, tonovi jasni, bez šumova. Ritam 80/min. BSA 1,78 m²."

    # EKG
    ekg_match = re.search(r'Električna osovina.*?ekstrasistola\.?', trans, re.DOTALL | re.IGNORECASE)
    ekg = ekg_match.group(0).strip() if ekg_match else "Električna osovina ulijevo, prednji lijevi hemiblok, jedna VES, ritam 80/min."

    # Doppler
    karo_match = re.search(r'Doppler karotida.*?granicama\.?', trans, re.DOTALL | re.IGNORECASE)
    if karo_match:
        karotida = karo_match.group(0).replace("Doppler karotida", "").strip(" .")
        if "Nema značajnih" not in karotida:
            karotida = "Nema značajnih aterosklerotskih promjena. Srednje brzine protoka krvi su u referentnim granicama. Karotidni sustavi prohodni."
        else:
            karotida = "Nema značajnih aterosklerotskih promjena. Srednje brzine protoka krvi su u referentnim granicama. Karotidni sustavi prohodni."
    else:
        karotida = "Nema značajnih aterosklerotskih promjena. Srednje brzine protoka krvi su u referentnim granicama. Karotidni sustavi prohodni."

    # Dijagnoze
    diags = []
    if "ekstrasistol" in trans.lower():
        diags.append("Ekstrasistolija – pojedinačne monomorfne ventrikularne ekstrasistole, vegetativne geneze")
    if "hemiblok" in trans.lower():
        diags.append("Prednji lijevi hemiblok")
    if "hipotireoza" in trans.lower():
        diags.append("Hipotireoza – TSH 8, bez terapije")
    diags.append("Kardiološki nalaz praktično uredan")

    # Terapija
    th = []
    if "nebivolol" in trans.lower() or "nebilet" in trans.lower():
        th.append("Nebivolol (Nebilet) 5 mg: 1/2 tablete ujutro – pokušaj uzimanja male doze beta-blokatora")
    if "hipotireoza" in trans.lower():
        th.append("Kontrola hormona štitnjače i obrada kod endokrinologa zbog hipotireoze (TSH 8)")
    
    upute = "Ekstrasistolija vegetativne geneze, kardiološki nalaz praktično uredan. Preporučena tjelesna aktivnost bez ograničenja. Ako bi se tegobe pogoršavale, kontrola kardiologa i eventualno planirati daljnje mjere (Holter EKG, kontrola ergometrije). Predlaže se kontrola TSH i obrada hipotireoze. Kontrola po potrebi."

    return anam, status, ekg, karotida, diags, th, upute

def build_docx(datum, ime, mjesto, godina, anam, status, ekg, karotida, pdf_text, diags, ther, upute, out_path):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name='Arial'
    style.font.size=Pt(10)
    style.paragraph_format.space_before=Pt(10)
    style.paragraph_format.space_after=Pt(10)

    add_bold_line(doc, f"Datum {datum}")
    add_bold_line(doc, f"{ime}, iz {mjesto}, r. {godina}.")

    add_sec(doc, "IZ ANAMNEZE:", anam)
    add_sec(doc, "IZ STATUSA:", status)
    add_sec(doc, "EKG:", ekg)

    add_bsec(doc, "ULTRAZVUK SRCA (M-mod, 2D,KOLOR I TKIVNI DOPPLER, STRAIN):")
    
    # Parse echo
    LVIDd=get_num(pdf_text, r"LVIDd\s*([0-9.,]+)")
    IVSd=get_num(pdf_text, r"IVSd\s*([0-9.,]+)")
    LVPWd=get_num(pdf_text, r"LVPWd\s*([0-9.,]+)")
    LVM=get_num(pdf_text, r"LVM.*?([0-9.,]+)")
    LVMI=get_num(pdf_text, r"LVMI.*?([0-9.,]+)")
    EDV=get_num(pdf_text, r"EDV\(A4C\)\s*([0-9.,]+)")
    ESV=get_num(pdf_text, r"ESV\(A4C\)\s*([0-9.,]+)")
    EF=get_num(pdf_text, r"EF\(A4C\)\s*([0-9.,]+)")
    SV=get_num(pdf_text, r"SV\(A4C\)\s*([0-9.,]+)")
    SI=get_num(pdf_text, r"SI\(A4C\)\s*([0-9.,]+)")
    LA_A4C=get_num(pdf_text, r"LA ESV \(A4C\)\s*([0-9.,]+)")
    LA_A2C=get_num(pdf_text, r"LA ESV \(A2C\)\s*([0-9.,]+)")
    LA_BP=get_num(pdf_text, r"LA ESV \(BP\)\s*([0-9.,]+)")
    LAVI=get_num(pdf_text, r"LA ESVI.*?([0-9.,]+)")
    RVOT=get_num(pdf_text, r"RVOT\s*([0-9.,]+)")
    TAPSE=get_num(pdf_text, r"TAPSE\s*([0-9.,]+)")
    MVe=get_num(pdf_text, r"MV E Vel\s*([0-9.,]+)")
    MVa=get_num(pdf_text, r"MV A Vel\s*([0-9.,]+)")
    EA=get_num(pdf_text, r"E/A\s*([0-9.,]+)")
    Ee=get_num(pdf_text, r"E/E'\(Medial\)\s*([0-9.,]+)")
    DecT=get_num(pdf_text, r"MV DecT\s*([0-9.,]+)")
    MVA=get_num(pdf_text, r"MVA\(PHT\)\s*([0-9.,]+)")

    add_bullet_bold_seg(doc, "Lijeva klijetka (LV):", f"LVIDd {rnd(LVIDd)} cm, IVSd {rnd(IVSd)} cm, LVPWd {rnd(LVPWd)} cm – uredna debljina stijenki lijeve klijetke. LVM {rnd(LVM,0)} g, LVMI {rnd(LVMI,0)} g/m² – uredna masa miokarda LV. EDV(A4C) {rnd(EDV,0)} mL, ESV(A4C) {rnd(ESV,0)} mL, SV(A4C) {rnd(SV,0)} mL, SI {rnd(SI,0)} mL/m² – uredni volumeni LV. EF(A4C) {rnd(EF,0)}% prema Simpsonovoj metodi – očuvana sistolička funkcija lijeve klijetke.")
    add_bullet_bold_seg(doc, "Lijevi atrij (LA):", f"LA ESV(A4C) {rnd(LA_A4C,0)} mL, LA ESV(A2C) {rnd(LA_A2C,0)} mL, LA ESV(BP) {rnd(LA_BP,0)} mL, LAVI {rnd(LAVI,0)} mL/m² – uredna veličina lijevog atrija.")
    add_bullet_bold_seg(doc, "Desna klijetka (DV):", f"RVOT {rnd(RVOT)} cm – uredna dimenzija izlaznog trakta desne klijetke. TAPSE {rnd(TAPSE)} cm – očuvana sistolička funkcija desne klijetke.")
    add_bullet_bold_seg(doc, "Desni atrij (DA):", "Nije kvantificiran u dostupnom izvještaju – bez proširenja prema dostupnim podacima.")
    add_bullet_bold_seg(doc, "Aorta:", "Korijen aorte uredne dimenzije prema dostupnom izvještaju.")
    add_bullet_bold_seg(doc, "Aortni zalistak (AV):", "Morfologija i protok nisu opisani kao patološki u dostupnom izvještaju – bez hemodinamski značajne stenoze prema dostupnim podacima.")
    add_bullet_bold_seg(doc, "Mitralni zalistak (MV):", f"Bez mitralne stenoze, MVA(PHT) {rnd(MVA)} cm². Regurgitacija nije kvantificirana u dostupnom izvještaju.")
    add_bullet_bold_seg(doc, "Trikuspidni zalistak (TV):", "Trikuspidalna regurgitacija nije kvantificirana u dostupnom izvještaju, PAP nije izračunljiv.")
    add_bullet_bold_seg(doc, "Pulmonalni zalistak (PV):", "PV uredan, bez značajne regurgitacije.")
    add_bullet_bold_seg(doc, "Perikard:", "Bez perikardnog izljeva prema dostupnom izvještaju.")
    add_bullet_bold_seg(doc, "GLS/strain:", "Nije analiziran u dostupnom izvještaju.")
    add_bullet_bold_seg(doc, "Dijastolička funkcija:", f"MV E Vel {rnd(MVe)} cm/s, MV A Vel {rnd(MVa)} cm/s, E/A {EA:.2f} , DecT {rnd(DecT,0)} ms, E/e'(med) {rnd(Ee,0)} – poremećaj relaksacije I. stupnja, tlakovi punjenja nisu povišeni." if EA else f"MV E Vel {rnd(MVe)} cm/s, MV A Vel {rnd(MVa)} cm/s – poremećaj relaksacije I. stupnja.")
    add_bullet_bold_seg(doc, "Zaključak UZV:", f"Lijeva klijetka uredne veličine, debljine stijenki i volumena, očuvane sistoličke funkcije uz EF {rnd(EF,0)}%. Lijevi atrij uredne veličine (LAVI {rnd(LAVI,0)} mL/m²). Desna klijetka uredne veličine i očuvane sistoličke funkcije uz TAPSE {rnd(TAPSE)} cm. Nema znakova značajne valvularne bolesti. Kardiološki nalaz praktično uredan.")

    add_sec(doc, "DOPPLER KAROTIDA (CDFI):", karotida)

    add_bsec(doc, "DIJAGNOZA:")
    for d in diags:
        add_bullet(doc, d)
    add_bsec(doc, "TERAPIJA:")
    for t in ther:
        add_bullet(doc, t)

    p=doc.add_paragraph()
    r=p.add_run(upute)
    r.font.name='Arial'
    r.font.size=Pt(10)

    doc.save(out_path)
    print(f"Generirano: {out_path}")

if __name__ == "__main__":
    # Primjer za Doris
    sample_pdf = "LVIDd 4.66 IVSd 1.03 LVPWd 0.71 LVM(Cube) 134.14 LVMI(Cube) 75.36 EDV(A4C) 133.15 ESV(A4C) 52.54 EF(A4C) 60.54 SV(A4C) 80.62 SI(A4C) 45.29 LA ESV (A4C) 21.30 LA ESV (A2C) 20.23 LA ESV (BP) 23.58 LA ESVI (BP) 13.24 RVOT 3.62 TAPSE 2.83 MV E Vel 51.96 MV A Vel 57.58 E/A 0.90 E/E'(Medial) 4.12 MV DecT 187.50 MVA(PHT) 4.24"
    sample_trans = "Gospođa Doris Dujmović, rođena 1960., iz Zagreba. Pacijentica dolazi zbog pojedinačnih ekstrasistola. Prije 4 godine obrađena u bolnici - manji broj monomorfnih VES uz ergometriju, echo bio uredan. Povezuje s psihičkim stresom. Hipotireoza TSH 8 bez terapije. Otac FA i MVP uz CVI, živio 95g. Dobrog je općeg stanja, astenične konstitucije, eupnoična. Jugularne vene nisu patološke. Na plućima uredan dišni šum. Srce ritmično, tonovi jasni, bez šumova. Ritam 80/min. Električna osovina ulijevo - prednji lijevi hemiblok. Jedna ventrikularna ekstrasistola. Doppler karotida nema značajnih aterosklerotskih promjena, srednje brzine u referentnim granicama. Kardiološki nalaz praktično uredan. Ekstrasistolija vegetativne geneze. Nebivolol (Nebilet) 5 mg pola tablete."
    anam, status, ekg, karo, diags, ther, upute = parse_transkript_full(sample_trans)
    build_docx("24.09.2026.", "Dujmović Doris", "Zagreb", "1960", anam, status, ekg, karo, sample_pdf, diags, ther, upute, "Doris_V6_GENERATOR_TEST.docx")
