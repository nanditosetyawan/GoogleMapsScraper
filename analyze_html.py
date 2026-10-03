from bs4 import BeautifulSoup
import re
import json

with open('logs/error_no_review_btn_20261003_194718.html', encoding='utf-8') as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

with open('analysis_output.txt', 'w', encoding='utf-8') as out:
    out.write("=== Buttons with text ===\n")
    buttons = soup.find_all('button')
    for b in buttons:
        text = b.get_text(strip=True)
        if text or b.get('aria-label'):
            out.write(f"Btn: '{text}' | aria-label='{b.get('aria-label')}' | class='{b.get('class')}' | jsaction='{b.get('jsaction')}'\n")

    out.write("\n=== Role=tab ===\n")
    tabs = soup.find_all(attrs={"role": "tab"})
    for tab in tabs:
        out.write(f"Tab text: {tab.get_text(strip=True)} | aria-label={tab.get('aria-label')} | class={tab.get('class')}\n")
    
    out.write("\n=== Divs with Ulasan or Review ===\n")
    for el in soup.find_all(string=re.compile('(?i)ulasan|review')):
        parent = el.parent
        out.write(f"Text: '{el.strip()}' | Parent: <{parent.name}> class={parent.get('class')} role={parent.get('role')} aria-label={parent.get('aria-label')}\n")

    out.write("\n=== ARIA-LABELS containing review/ulasan/bintang ===\n")
    for el in soup.find_all(attrs={"aria-label": re.compile('(?i)ulasan|review|bintang|penilaian')}):
        out.write(f"<{el.name}> aria-label='{el.get('aria-label')}' class='{el.get('class')}' text='{el.get_text(strip=True)}'\n")

print("Analysis written to analysis_output.txt")
