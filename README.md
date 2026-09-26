# Kardiološki Generator - Online V6 Konačna

**Konačna verzija po predlošku `PREDLOZAK-KARDIOLOSKOG-NALAZA_bold_bullet.docx` - Arial 10 - Bold anatomski**

Online generator za strukturirane kardiološke nalaze na hrvatskom jeziku.

## ✅ Značajke V6 Konačna

- **Bold naslovi sekcija:** IZ ANAMNEZE, IZ STATUSA, EKG, ULTRAZVUK SRCA (M-mod, 2D,KOLOR I TKIVNI DOPPLER, STRAIN), DOPPLER KAROTIDA (CDFI), DIJAGNOZA, TERAPIJA
- **Bold anatomski segmenti:** LV, LA, DV, DA, aorta, AV, MV, TV, PV, perikard, GLS/strain, Zaključak UZV
- **Font:** Arial 10, bullet `•`, 1 prazan red između sekcija, 0 praznih redova između bulleta
- **Rich ehokardiografija:** LVIDd, IVSd, LVPWd, LVM, LVMI, EDV, ESV, SV, SI, EF, LA ESV, LAVI, RVOT, TAPSE, MV E/A, E/e', DecT, MVA, AV Vmax/PGmax/PGmean, itd.
- **Puni sadržaj:** Doppler karotida puni opis, dijagnoza 4 stavke, terapija 2 stavke, upute

## 🚀 Brzi start - Online verzija (bez instalacije)

### Opcija 1: GitHub Pages (preporučeno za trajno online)

1. Forkaj ovaj repo ili napravi novi repo `kardioloski-generator-v6`
2. Uploadaj sve fileove
3. Idi na **Settings → Pages → Source: Deploy from a branch → main / root**
4. Nakon 1 minute dobiješ link: `https://tvojeime.github.io/kardioloski-generator-v6/`
5. Otvori `OFFLINE_V6_KONACNA_PUNA_BOLD.html` ili `index.html` - radi odmah!

### Opcija 2: Dvoklik offline

- Windows: dvoklik na `OFFLINE_V6_KONACNA_PUNA_BOLD.html`
- Ne treba Python, ne treba internet (osim za prvi load pdf.js i docx.js CDN-a)

### Opcija 3: Python

```bash
pip install python-docx
python offline_V6_generator.py
```

## 📁 Struktura projekta

```
kardioloski-generator-v6/
├── index.html                              # ONLINE V6 - glavna stranica za GitHub Pages
├── OFFLINE_V6_KONACNA_PUNA_BOLD.html       # Offline HTML dvoklik verzija
├── offline_V6_generator.py                 # Python generator
├── Pokreni_V6.bat                          # Windows batch starter
├── PREDLOZAK-KARDIOLOSKOG-NALAZA_bold_bullet.docx  # Originalni predložak s pravilima
├── PRIMJER_Doris_V6_KONACNA.docx           # Primjer ispravnog nalaza (Doris Dujmović)
├── PRIMJER_Juric_Ivo.docx                  # Primjer Jurić Ivo
└── README.md                               # Ova datoteka
```

## 🔧 Kako koristiti

1. **Transkript:** Zalijepi tekst diktafona (anamneza, status, EKG, Doppler, dijagnoza, terapija, upute)
2. **Echo PDF:** Uploadaj sirovi izvještaj s ultrazvučnog aparata (CARD Report)
3. **Podaci pacijenta:** Datum, ime, mjesto, godina rođenja
4. **Klik GENERIRAJ:** Dobiješ DOCX Arial 10, boldano po pravilima, spreman za print

## 📋 Pravila iz predloška

Iz `PREDLOZAK-KARDIOLOSKOG-NALAZA_bold_bullet.docx`:

- Datum i ime pacijenta: cijeli red boldano
- Naslovi sekcija boldano, sadržaj u istom retku (osim UZV, DIJAGNOZA, TERAPIJA gdje bullet ide u sljedeći red)
- UZV sekcija: grupiranje po anatomskim segmentima, svaki segment boldano + vrijednosti + interpretacija
- Bullet oznaka: točno `•`, bez crtica ili zvjezdica
- Lijekovi: generički (tvornički) jačina: doza način

## 🔄 Verzijska povijest

- **V1:** Osnovni regex parser - krnji opis, LA n/a, AV n/a
- **V2:** Rich ehokardiogram - LVM, EDV, LAVI, TAPSE - ali samo UZV, bez sekcija
- **V3:** PUNE SEKCIJE - anamneza, status, EKG, Doppler, dijagnoza, terapija
- **V4:** Bold fix - bold naslovi IZ ANAMNEZE, IZ STATUSA, EKG, DOPPLER KAROTIDA
- **V5:** Bold anatomski segmenti LV, LA, DV, DA, aorta, AV, MV, TV, PV, perikard, GLS - ali krnji sadržaj (bug)
- **V6 KONAČNA:** Spoj V4 (puni sadržaj) + V5 (bold anatomski) = 100% po predlošku

## 🌐 Online sačuvati projekt

### Perplexity Spaces

1. https://www.perplexity.ai/spaces → New Space
2. Name: `Kardiološki Generator V6`
3. Instructions: kopiraj pravila iz PREDLOZAK docx
4. Knowledge: uploadaj sve fileove iz ovog repoa

### Meta AI

Ovaj chat je tvoj projekt - bookmarkaj ga. Generator radi unutar chata kao artifact.

## 📄 Licenca

Interni projekt - za kliničku uporabu.

## 👨‍⚕️ Autor

Kardiološki projekt - V6 Konačna - 2026.

---

**Napomena:** Ova verzija je 100% po predlošku bold_bullet.docx - provjereno: bold LV, LA, DV, DA, aorta, AV, MV, TV, PV, perikard, GLS, Zaključak UZV + puni Doppler + puna dijagnoza/terapija + Arial 10.
