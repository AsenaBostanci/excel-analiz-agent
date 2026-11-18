import datetime as dt
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

YEAR = 2026
MONTH = 11
FIRST_DAY, LAST_DAY = 1, 12
OUTPUT = "usd_kurlari_1-12_kasim_2026.xlsx"
URL = "https://www.tcmb.gov.tr/kurlar/{ym}/{dmy}.xml"


def fetch_usd(day: dt.date):
    url = URL.format(ym=day.strftime("%Y%m"), dmy=day.strftime("%d%m%Y"))
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            root = ET.fromstring(resp.read())
    except (urllib.error.HTTPError, ET.ParseError):
        return None
    for cur in root.findall("Currency"):
        if cur.get("Kod") == "USD":
            get = lambda tag: float(cur.findtext(tag))
            return (get("ForexBuying"), get("ForexSelling"),
                    get("BanknoteBuying"), get("BanknoteSelling"))
    return None


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "USD Kurları"
    headers = ["Tarih", "Gün", "Döviz Alış", "Döviz Satış",
               "Efektif Alış", "Efektif Satış", "Not"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")

    gunler = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
    for d in range(FIRST_DAY, LAST_DAY + 1):
        day = dt.date(YEAR, MONTH, d)
        rates = fetch_usd(day)
        if rates:
            ws.append([day, gunler[day.weekday()], *rates, ""])
        else:
            ws.append([day, gunler[day.weekday()], None, None, None, None,
                       "Kur ilan edilmedi (hafta sonu/tatil)"])
        ws.cell(ws.max_row, 1).number_format = "DD.MM.YYYY"
        for col in range(3, 7):
            ws.cell(ws.max_row, col).number_format = "0.0000"

    for i, w in enumerate([12, 12, 14, 14, 14, 14, 38], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    wb.save(OUTPUT)
    print("Kaydedildi:", OUTPUT)


if __name__ == "__main__":
    main()
