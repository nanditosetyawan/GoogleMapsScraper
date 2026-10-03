import time
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

URL = "https://maps.app.goo.gl/p5ybQRvDtJxicFCs9"

def run_diagnostic():
    print("Mulai diagnostic script...")
    out_dir = Path("diagnostic_output")
    out_dir.mkdir(exist_ok=True)
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        # Menggunakan viewport standar desktop Chrome (1920x1080)
        context = browser.new_context(viewport={"width": 1920, "height": 1080}, locale="id-ID")
        page = context.new_page()
        
        print(f"Membuka URL: {URL}")
        page.goto(URL, wait_until="networkidle", timeout=60000)
        
        # Waktu ke-5 detik
        print("Menunggu 5 detik...")
        time.sleep(5)
        page.screenshot(path=str(out_dir / "screenshot_05s.png"))
        
        # Waktu ke-10 detik
        print("Menunggu 10 detik...")
        time.sleep(5)
        page.screenshot(path=str(out_dir / "screenshot_10s.png"))
        
        # Waktu ke-20 detik
        print("Menunggu 20 detik...")
        time.sleep(10)
        page.screenshot(path=str(out_dir / "screenshot_20s.png"))
        
        print("Mengumpulkan data DOM...")
        data = {
            "url": page.url,
            "title": page.title(),
            "body_text": page.locator("body").inner_text(),
        }
        
        # Ambil semua button
        buttons = page.locator("button").all()
        data["buttons"] = [
            {
                "aria-label": b.get_attribute("aria-label"),
                "class": b.get_attribute("class"),
                "text": b.inner_text().strip(),
            } for b in buttons
        ]
        
        # Ambil semua a[href]
        links = page.locator("a[href]").all()
        data["links"] = [
            {
                "href": a.get_attribute("href"),
                "text": a.inner_text().strip(),
            } for a in links
        ]
        
        # Elemen dengan kata ulasan/reviews
        ulasan_els = page.locator("text=/ulasan|review|129/i").all()
        data["ulasan_elements"] = [
            {
                "tag": e.evaluate("el => el.tagName"),
                "text": e.inner_text().strip(),
                "class": e.get_attribute("class")
            } for e in ulasan_els if e.is_visible()
        ]
        
        # Simpan data ke JSON
        with open(out_dir / "dom_data.json", "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
            
        print(f"Data disimpan di {out_dir}/dom_data.json")
        browser.close()

if __name__ == "__main__":
    run_diagnostic()
