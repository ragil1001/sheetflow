import json
import random
import re
import sys
from contextlib import suppress
from decimal import Decimal, InvalidOperation
from math import ceil
from pathlib import Path
from time import monotonic, sleep
from urllib.parse import quote
from camoufox.fingerprints import from_fpgen, generate_fingerprint
from camoufox.sync_api import Camoufox
from camoufox.utils import get_screen_cons
from openpyxl import load_workbook
from playwright.sync_api import Error, expect

TIMEOUT = 30_000
TIMEOUT_CARI_USAHA = 5_000
PESAN_BOT = re.compile(r"\bBOT-\d+\b|\bbot\s+detected\b|(?:terdeteksi|teridentifikasi)\s+(?:sebagai\s+)?(?:bot|robot)\b|koneksi\s+anda\s+sebagai\s+bot\b|(?:kami\s+mendeteksi|sistem\s+kami\s+mendeteksi)\s+(?:perilaku\s+yang\s+tidak\s+wajar|koneksi)|perilaku\s+yang\s+tidak\s+wajar\s+pada\s+koneksi", re.I)

FIELDS = {
    "tk_laki_laki": "Pekerja laki-laki",
    "tk_perempuan": "Pekerja perempuan",
    "tk_dibayar": "Pekerja dibayar",
    "tk_tidak_dibayar": "Pekerja tidak dibayar",
    "upah_gaji": "Total upah dan gaji, serta jaminan sosial pegawai",
    "perbaikan_produksi": "Biaya produksi",
    "perbaikan_biaya_pembelian": "Biaya pembelian barang yang terjual",
    "perbaikan_operasional": "Biaya operasional",
    "perbaikan_nonoperasional": "Biaya non-operasional",
    "perbaikan_pendapatan": "Nilai pendapatan barang dan jasa",
}
SELECTORS = {
    "revoke": 'button:visible:has(svg.tabler-icon-rotate-clockwise)',
    "konfirmasi_revoke": '[role="alertdialog"]:visible:has-text("Revoke Assignment?") button:text-is("Konfirmasi")',
    "edit": 'button:visible:has(svg.tabler-icon-edit)',
    "blok_ii": '.fasih-form-sidebar:visible > div[title="SE2026 - L BLOK II"]:visible',
    "kartu_usaha": '#se2026_nested [data-nested-view="true"]:visible',
    "nama_usaha": '[id^="nama_usaha_edit#"]:visible',
    "catatan": '.fasih-form-sidebar:visible > div[title="CATATAN"]:visible',
    "switch": '#cek_anomali_button input[role="switch"]',
    "kontrol_switch": '#cek_anomali_button div[id$="-control"]',
    "modal": r'div.tw\:border\!.tw\:relative.tw\:flex.tw\:w-full.tw\:flex-col.tw\:gap-4.tw\:bg-background.tw\:p-6.tw\:pt-4.tw\:shadow-lg\!.tw\:sm\:rounded-lg:visible',
}

def pola_usaha(target):
    target = " ".join(target.split())
    parts = re.fullmatch(r"(.+?)\s*(?:\(([^()<>]+)\)|<([^()<>]+)>?)", target)
    text = re.escape(target).replace(r"\ ", r"\s+")
    if parts:
        owner = re.escape((parts[2] or parts[3]).strip()).replace(r"\ ", r"\s+")
        text = re.escape(parts[1].strip()).replace(r"\ ", r"\s+") + r"(?:[\s.,/\-]+[^()<>]*)?\s*(?:\(\s*" + owner + r"\s*\)|<\s*" + owner + r"\s*>?)"
    return re.compile(r"^\s*" + text + r"\s*$", re.I)

def inputan(page, label):
    text = r"[\W_]+".join(re.escape(w) for w in re.findall(r"\w+", label))
    if label == FIELDS["perbaikan_pendapatan"]:
        text = r"Nilai\W+(?:(?:produksi|pendapatan|penjualan)\W+){1,3}barang\W+dan\W+jasa"
    prefix = r"(?:(?!\b(?:total|rincian|jumlah|subtotal|petunjuk|contoh|catatan|keterangan|akumulasi)\b)[\s\S])*?"
    pattern = re.compile(r"^" + prefix + r"\b(?:" + text + r")\b[\s\S]*$", re.I)
    inputs = 'input[not(@type) or @type="text" or @type="number" or @type="tel"]'
    rows = page.get_by_text(pattern).filter(visible=True).locator(
        "xpath=ancestor-or-self::div[contains(concat(' ', normalize-space(@class), ' '), ' tw:rounded-lg ') "
        "and contains(concat(' ', normalize-space(@class), ' '), ' tw:gap-2 ')][1]")
    field = page.get_by_label(pattern).and_(page.locator("xpath=//" + inputs))
    return field.or_(rows.locator(f"xpath=self::*[count(.//{inputs})=1]//{inputs}")).filter(visible=True)

def buka_usaha(page, target):
    target = " ".join(target.split())
    cards = page.locator(SELECTORS["kartu_usaha"])
    expect(cards.first, "Daftar kartu usaha BLOK II tidak muncul").to_be_visible(timeout=TIMEOUT)
    daftar = page.locator("#se2026_nested")
    pola_ringkas = pola_usaha(re.sub(r"\s*(?:\([^()<>]+\)|<[^()<>]+>?)$", "", target))
    kandidat = daftar.get_by_text(pola_usaha(target)).or_(daftar.get_by_text(pola_ringkas)).filter(visible=True)
    # ponytail: kartu muncul menandai daftar siap; gunakan indikator selesai jika tersedia.
    terlihat = " | ".join(" ".join(text.split()) for text in cards.all_inner_texts())[:1500]
    try:
        expect(kandidat.first, "Kartu usaha tidak ditemukan: " + target + " | Kartu terlihat: " + terlihat).to_be_visible(timeout=TIMEOUT_CARI_USAHA)
    except AssertionError as error:
        if page.is_closed():
            raise
        raise ValueError("Kartu usaha tidak ditemukan: " + target + " | Kartu terlihat: " + terlihat) from error
    pattern = re.compile(r"^\s*" + re.escape(target).replace(r"\ ", r"\s+") + r"\s*$", re.I)
    name = daftar.get_by_text(pattern).filter(visible=True)
    if not name.count():
        pattern = pola_usaha(target)
        name = daftar.get_by_text(pattern).filter(visible=True)
    ringkas = not name.count()
    if ringkas:
        name = daftar.get_by_text(pola_ringkas).filter(visible=True)
    card = name.locator('xpath=ancestor-or-self::*[@data-nested-view="true"][1]')
    assert card.count() <= 1, "Nama usaha cocok dengan beberapa kartu; tidak memilih sembarang kartu"
    expect(card, "Kartu usaha tidak ditemukan: " + target).to_have_count(1, timeout=TIMEOUT_CARI_USAHA)
    name.first.click()
    nama = page.locator(SELECTORS["nama_usaha"])
    try:
        expect(nama).to_have_count(1, timeout=TIMEOUT)
        field = nama.locator("input:visible").or_(nama.and_(page.locator("input:visible")))
        if field.count():
            expect(field).to_have_count(1, timeout=TIMEOUT)
            expect(field).not_to_have_value(re.compile(r"^\s*$"), timeout=TIMEOUT)
            terlihat = " ".join(field.input_value().split())
            cocok = pola_usaha(target).fullmatch(terlihat)
        else:
            terlihat = " ".join(nama.inner_text().split())
            cocok = nama.get_by_text(pola_usaha(target)).count() == 1
    except AssertionError as error:
        if page.is_closed():
            raise
        raise ValueError("Nama usaha dalam form tidak dapat diverifikasi: " + target) from error
    if not cocok:
        raise ValueError("Nama usaha dalam form tidak sesuai target Excel: " + target + " | Nama form: " + terlihat)

def isi_remarks(page, field, text):
    wadah = field.locator('xpath=ancestor::*[contains(@id, "#")][1]')
    tombol = wadah.locator(r'button.tw\:size-8.tw\:relative.tw\:text-primary.tw\:\@lg\/form\:flex\!.tw\:hidden:visible:has(> svg.tw\:size-4)')
    expect(tombol, "Tombol catatan pada input yang diubah tidak ditemukan/unik").to_have_count(1, timeout=TIMEOUT)
    tombol.click()
    editor = page.get_by_placeholder("Tambah catatan di sini...", exact=True).filter(visible=True)
    editor.fill(text)
    expect(editor).to_have_value(text, timeout=TIMEOUT)
    page.get_by_role("button", name="Simpan", exact=True).click()
    page.get_by_role("button", name="Dismiss", exact=True).click()
    expect(editor).to_be_hidden(timeout=TIMEOUT)

def angka(value):
    if value is None or not str(value).strip():
        return None
    text = "0" if str(value).strip() in ("-", "–", "—") else str(value).strip()
    if isinstance(value, str):
        text = text.replace(".", "").replace(",", ".")
    number = Decimal(text)
    if isinstance(value, bool) or not number.is_finite() or number < 0:
        raise ValueError("Angka tidak valid")
    return number

def simpan(book, path):
    temp = path.with_suffix(".tmp.xlsx")
    book.save(temp)
    temp.replace(path)

def identitas(path):
    if path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
        assert isinstance(data, dict) and isinstance(data.get("navigator"), dict), "Fingerprint tersimpan tidak valid"
        assert "Firefox/" in data["navigator"].get("userAgent", ""), "Fingerprint harus Firefox"
        return data
    data = generate_fingerprint(os="windows", screen=get_screen_cons())
    data.setdefault("window", {})["screenY"] = from_fpgen(data)["window.screenY"]
    path.parent.mkdir(exist_ok=True)
    temp = path.with_suffix(".tmp.json")
    temp.write_text(json.dumps(data), encoding="utf-8")
    temp.replace(path)
    return data

def tunggu_kirim(page, tujuan):
    gagal = page.get_by_text(re.compile(r"\bpre[\s_-]?signed[\s_-]*url\b", re.I)).filter(visible=True).first
    expect(tujuan.filter(visible=True).or_(gagal).first).to_be_visible(timeout=TIMEOUT)
    assert not gagal.is_visible(), "Kirim gagal: FASIH menampilkan kesalahan presigned URL; hasil kirim belum dipastikan"

def pesan_galat(page):
    modal = page.locator(SELECTORS["modal"])
    tunggu_kirim(page, modal)
    tombol = modal.get_by_role("button", name=re.compile(r"^Galat\s+[1-9]\d*\s+Perlu diperbaiki$"))
    if not tombol.count():
        return ""
    tombol.click()
    rincian = page.locator(SELECTORS["modal"]).filter(has=page.get_by_text(re.compile(r"^Total:\s*\d+\s*/\s*\d+$")))
    pesan = rincian.locator(r"ul.tw\:text-destructive > li")
    expect(pesan.first).to_be_visible(timeout=TIMEOUT)
    lines = []
    for item in pesan.all():
        wadah = item.locator("xpath=../..")
        label = wadah.locator("[title]")
        judul = label.first.get_attribute("title") if label.count() else " ".join(wadah.locator("xpath=./div").all_inner_texts())
        lines.append((judul.strip() + " — " if judul else "") + item.inner_text().strip())
    text = "\n".join(lines)
    assert text and len(text) <= 32767, "Pesan Galat kosong atau melebihi batas sel Excel"
    return text

def sesi(browser, page, path, simpan=False):
    if simpan:
        tab = json.loads(path.read_text(encoding="utf-8"))["tab"] if path.exists() else {}
        if not page.is_closed() and page.url.startswith("https://fasih-sm.bps.go.id/"):
            tab = page.evaluate("() => ({...sessionStorage})")
        data = {"browser": browser.storage_state(indexed_db=True), "tab": tab}
        temp = path.with_suffix(".tmp.json")
        temp.write_text(json.dumps(data), encoding="utf-8")
        temp.replace(path)
    elif path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
        browser.set_storage_state(data["browser"])
        page.add_init_script("if (location.origin === 'https://fasih-sm.bps.go.id' && !sessionStorage.length) "
                             "for (const [k,v] of Object.entries(" + json.dumps(data["tab"]) + ")) sessionStorage.setItem(k,v);")

def tab_survei(browser):
    browser.cookies()
    return next(p for p in reversed(browser.pages) if not p.is_closed() and p.url.startswith("https://fasih-sm.bps.go.id/app/surveys"))

def gangguan(page):
    bot = page.get_by_text(PESAN_BOT).and_(page.locator(':not(#se2026_nested [data-nested-view="true"], #se2026_nested [data-nested-view="true"] *, [id^="nama_usaha_edit#"], [id^="nama_usaha_edit#"] *)'))
    akses = page.get_by_text("Gagal Mendapatkan Akses Approve/Reject/Edit Assignment", exact=True)
    return bot.or_(akses).or_(page.get_by_role("button", name="Refresh Page", exact=True)).filter(visible=True).first

def pasang_refresh(page):
    pesan = gangguan(page)
    def ulang():
        print("  [PULIH] BOT/akses gagal/Refresh Page; menutup halaman untuk restart browser.", flush=True)
        if len(page.context.pages) == 1:
            page.context.new_page()
        page.close()
    page.remove_locator_handler(pesan)
    page.add_locator_handler(pesan, ulang, times=1, no_wait_after=True)

def jeda(cfg, sisa):
    sisa -= 1
    if sisa == 0:
        return random.uniform(cfg["rest_min"], cfg["rest_max"]) * 60000, random.randint(cfg["batch_min"], cfg["batch_max"])
    return random.uniform(cfg["delay_min"], cfg["delay_max"]) * 1000, sisa

def hitung_mundur(page, waktu, label="Istirahat batch"):
    akhir = monotonic() + waktu / 1000
    while True:
        tersisa = max(0, akhir - monotonic())
        menit, detik = divmod(ceil(tersisa), 60)
        print(f"\r{label}: {menit:02d}:{detik:02d}", end="", flush=True)
        if tersisa == 0:
            print(" - lanjut.")
            return
        if page is None:
            sleep(min(1, tersisa))
        else:
            page.wait_for_timeout(min(1000, tersisa * 1000))

def jalankan(pemulihan=False):
    root = Path(__file__).resolve().parent
    cfg = json.loads((root / "config.json").read_text(encoding="utf-8"))
    for key in ("delay", "batch", "rest"):
        low, high = cfg[key + "_min"], cfg[key + "_max"]
        if type(low) not in (int, float) or type(high) not in (int, float) or not 0 <= low <= high < float("inf"):
            raise ValueError("Rentang config tidak valid: " + key)
    if any(type(cfg[k]) is not int or cfg[k] < 1 for k in ("batch_min", "batch_max")):
        raise ValueError("Batch harus bilangan bulat positif")
    path = root / cfg["excel_file"]
    book = load_workbook(path)
    sheet = book[cfg["sheet_name"]]
    headers = [str(c.value).strip() for c in sheet[1]]
    headers = ["remarks" if h == "remark" else h for h in headers]
    assert all(headers.count(k) == 1 for k in ("id_fasih", "target_usaha", *FIELDS))
    assert all(headers.count(k) <= 1 for k in ("status", "keterangan", "catatan", "remarks"))
    rows = []
    for number, cells in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
        row = dict(zip(headers, cells))
        if str(row.get("status") or "").strip().upper() not in ("", "TODO", "RETRY"):
            continue
        assert all(isinstance(row[k], str) and row[k].strip() for k in ("id_fasih", "target_usaha"))
        assert row.get("remarks") is None or isinstance(row["remarks"], str), "Remarks harus berupa teks"
        values = {k: angka(row[k]) for k in FIELDS}
        assert all(v is None or v == v.to_integral_value() for k, v in values.items() if k.startswith("tk_"))
        rows.append((number, row, values))
    if "--cek-excel" in sys.argv or not rows:
        if not rows and "--cek-excel" not in sys.argv:
            print("Tidak ada baris kosong/TODO/RETRY untuk diproses.", flush=True)
        book.close()
        return
    status = headers.index("status") + 1 if "status" in headers else sheet.max_column + 1
    sheet.cell(1, status, "status")
    keterangan = headers.index("keterangan") + 1 if "keterangan" in headers else sheet.max_column + 1
    sheet.cell(1, keterangan, "keterangan")
    catatan = headers.index("catatan") + 1 if "catatan" in headers else sheet.max_column + 1
    sheet.cell(1, catatan, "catatan")
    with path.open("r+b"):
        pass
    fingerprint = identitas(root / ".fasih-profile" / "fingerprint.json")
    headless = False
    print(f"[MULAI] {len(rows)} baris; headless={headless}. Membuka browser dan memuat sesi tersimpan.", flush=True)
    with Camoufox(persistent_context=True, user_data_dir=str(root / ".fasih-profile"), fingerprint=fingerprint, i_know_what_im_doing=True, headless=headless, humanize=True, os="windows", locale="id-ID", timeout=TIMEOUT) as browser:
        browser.set_default_timeout(TIMEOUT)
        browser.set_default_navigation_timeout(TIMEOUT)
        def tab_baru(tab):
            try:
                if len(browser.pages) > 2:
                    tab.close()
                elif not tab.is_closed():
                    pasang_refresh(tab)
            except Error:
                if not tab.is_closed():
                    raise
        browser.on("page", tab_baru)
        for tab in browser.pages:
            pasang_refresh(tab)
        page = browser.pages[0] if browser.pages else browser.new_page()
        for tab in browser.pages:
            if tab != page:
                tab.close()
        sesi(browser, page, root / ".fasih-profile" / "sesi.json")
        print("[SESI] Memeriksa akses halaman Survei...", flush=True)
        try:
            response = page.goto("https://fasih-sm.bps.go.id/app/surveys?page=0&perPage=10&layout=list", wait_until="commit")
            if response and response.status >= 500:
                raise Error(f"Server HTTP {response.status}")
            expect(page.locator("body")).to_be_visible(timeout=TIMEOUT)
            if pemulihan or headless:
                page.wait_for_url("**/app/surveys?*", wait_until="domcontentloaded", timeout=TIMEOUT)
            if not pemulihan:
                input("Pemeriksaan awal selesai; tekan Enter untuk mulai... " if headless else "Setelah halaman Survei terbuka, tekan Enter di terminal... ")
            page = tab_survei(browser)
            page.wait_for_url("**/app/surveys?*", wait_until="commit")
            sesi(browser, page, root / ".fasih-profile" / "sesi.json", simpan=True)
        except (Error, AssertionError, StopIteration):
            print("[SESI] Survei belum siap; browser akan dibuka ulang otomatis.", flush=True)
            with suppress(Error):
                sesi(browser, page, root / ".fasih-profile" / "sesi.json", simpan=True)
            browser.remove_listener("page", tab_baru)
            book.close()
            return True
        print("[SESI] Halaman Survei terbuka. Mulai memeriksa assignment.", flush=True)
        hasil = dict.fromkeys(("SUCCESS", "NO_CHANGE", "ERROR"), 0)
        ulang_browser = False
        sisa = random.randint(cfg["batch_min"], cfg["batch_max"])
        for index, (number, row, values) in enumerate(rows):
            print(f"[{index + 1}/{len(rows)}] Baris Excel {number}: {row['target_usaha']}", flush=True)
            sheet.cell(number, status, "ERROR")
            simpan(book, path)
            assignment = page
            result, note = "ERROR", ""
            ulang = konfirmasi = False
            kirim_ulang = "Kirim belum selesai; pemeriksaan otomatis" in str(row.get("catatan") or "")
            tahap = "Membuka assignment detail"
            print("  " + tahap, flush=True)
            try:
                for tab in browser.pages:
                    if tab != assignment:
                        tab.close()
                pasang_refresh(page)
                page.goto("https://fasih-sm.bps.go.id/app/assignment-detail/" + quote(row["id_fasih"], safe=""), wait_until="commit")
                tahap = "Membuka tab Review"
                print("  " + tahap, flush=True)
                review = assignment.get_by_role("button", name="Review", exact=True)
                expect(review).to_be_visible(timeout=TIMEOUT)
                with assignment.expect_popup(timeout=TIMEOUT) as popup:
                    review.click(no_wait_after=True)
                page = popup.value
                tahap = "Menunggu Ringkasan"
                print("  " + tahap, flush=True)
                expect(page.get_by_role("button", name="Ringkasan", exact=True)).to_be_visible(timeout=TIMEOUT)
                tahap = "Menunggu Revoke/Edit"
                print("  " + tahap, flush=True)
                revoke, edit = page.locator(SELECTORS["revoke"]), page.locator(SELECTORS["edit"])
                expect(revoke.or_(edit).first).to_be_visible(timeout=TIMEOUT)
                changed = revoke.is_visible()
                if changed and kirim_ulang:
                    tahap = "Memeriksa hasil Kirim sebelumnya"
                    print("  " + tahap, flush=True)
                    page.locator(SELECTORS["blok_ii"]).click()
                    buka_usaha(page, row["target_usaha"])
                    sesuai = True
                    for column, label in FIELDS.items():
                        if values[column] is not None:
                            field = inputan(page, label)
                            assert field.count() == 1, "Input tidak terlihat/unik setelah nama usaha muncul: " + label
                            sesuai = sesuai and angka(field.input_value()) == values[column]
                    page.locator(SELECTORS["catatan"]).click()
                    if sesuai and page.locator(SELECTORS["switch"]).get_attribute("aria-checked") == "true":
                        result = "SUCCESS"
                        note = "Kirim sebelumnya terverifikasi: assignment sudah disubmit dan nilai sesuai Excel."
                if result != "SUCCESS":
                    if changed:
                        tahap = "Revoke dan konfirmasi"
                        revoke.click()
                        page.locator(SELECTORS["konfirmasi_revoke"]).click()
                    tahap = "Membuka Edit"
                    print("  " + tahap, flush=True)
                    edit.click(no_wait_after=True)
                    expect(edit).to_be_hidden(timeout=TIMEOUT)
                    tahap = "Membuka BLOK II"
                    page.locator(SELECTORS["blok_ii"]).click()
                    tahap = "BLOK II: membuka usaha " + row["target_usaha"]
                    print("  " + tahap, flush=True)
                    buka_usaha(page, row["target_usaha"])
                    dokumen = page.evaluate("performance.timeOrigin")
                    for column, label in FIELDS.items():
                        if values[column] is None:
                            continue
                        tahap = "Mencari/mengisi " + label
                        print("  " + tahap, flush=True)
                        field = inputan(page, label)
                        # ponytail: nama usaha menandai form siap; tunggu indikator selesai bila form kelak dirender bertahap.
                        assert field.count() == 1, "Input tidak terlihat/unik setelah nama usaha muncul: " + label
                        berbeda = not field.input_value().strip() or angka(field.input_value()) != values[column]
                        if berbeda:
                            field.fill(format(values[column], "f"))
                            field.press("Tab")
                            changed = True
                        assert angka(field.input_value()) == values[column], "Nilai web tidak sesuai Excel"
                        if berbeda and (row.get("remarks") or "").strip():
                            tahap = "Menyimpan remarks pada " + label
                            print("  " + tahap, flush=True)
                            isi_remarks(page, field, row["remarks"])
                    tahap = "CATATAN: memeriksa switch anomali"
                    print("  " + tahap, flush=True)
                    page.locator(SELECTORS["catatan"]).click()
                    switch = page.locator(SELECTORS["switch"])
                    expect(switch).to_have_count(1, timeout=TIMEOUT)
                    if switch.get_attribute("aria-checked") == "false":
                        page.locator(SELECTORS["kontrol_switch"]).click()
                        changed = True
                    expect(switch).to_have_attribute("aria-checked", "true", timeout=TIMEOUT)
                    assert page.evaluate("performance.timeOrigin") == dokumen, "Halaman dimuat ulang saat mengedit; periksa ulang perubahan sebelum retry"
                    if changed or kirim_ulang:
                        tahap = "Kirim: membuka dan membaca validasi"
                        print("  " + tahap, flush=True)
                        page.get_by_role("button", name="Kirim", exact=True).click()
                        detail = pesan_galat(page)
                        if detail:
                            cell = sheet.cell(number, keterangan, detail)
                            cell.data_type = "s"
                            note = "Galat validasi Kirim: " + detail
                        else:
                            assert page.evaluate("performance.timeOrigin") == dokumen, "Halaman dimuat ulang saat validasi; periksa ulang perubahan sebelum retry"
                            tahap = "Kirim: menunggu konfirmasi"
                            page.locator(SELECTORS["modal"]).get_by_role("button", name="Kirim", exact=True).click()
                            tunggu_kirim(page, page.get_by_text("Konfirmasi Kirim", exact=True))
                            assert page.evaluate("performance.timeOrigin") == dokumen, "Halaman dimuat ulang sebelum konfirmasi; hasil kirim perlu diperiksa sebelum retry"
                            tahap = "Konfirmasi Kirim; hasil kirim belum dipastikan"
                            print("  " + tahap, flush=True)
                            konfirmasi = True
                            page.get_by_role("button", name="Konfirmasi", exact=True).click()
                            tunggu_kirim(page, page.get_by_text("Assignment Berhasil DisubmitPerubahan Anda berhasil disimpan.", exact=True))
                            assert page.evaluate("performance.timeOrigin") == dokumen, "Halaman dimuat ulang saat konfirmasi; hasil kirim perlu diperiksa sebelum retry"
                            result = "SUCCESS"
                            tahap = "Kirim berhasil; menutup konfirmasi sukses"
                            page.get_by_role("button", name="Tetap di Halaman", exact=True).click()
                    else:
                        result = "NO_CHANGE"
                tahap = "Menyimpan sesi setelah " + result
                sesi(browser, page, root / ".fasih-profile" / "sesi.json", simpan=True)
            except (Error, AssertionError, ValueError, InvalidOperation) as error:
                error_text = str(error).strip()
                note = f"{tahap} | {type(error).__name__}: {error_text}"
                ulang = result == "ERROR" and (page.is_closed() or assignment.is_closed() or isinstance(error, Error) or "timeout" in error_text.lower() or "Halaman dimuat ulang" in error_text or "presigned URL" in error_text or bool(PESAN_BOT.search(error_text)))
                with suppress(Error):
                    ulang = ulang or any(not tab.is_closed() and gangguan(tab).is_visible() for tab in (page, assignment))
                ulang = ulang and not isinstance(error, (ValueError, InvalidOperation))
                if result == "ERROR" and (konfirmasi or (ulang and tahap.startswith("Kirim:"))):
                    ulang = True
                    note = tahap + " | Kirim belum selesai; pemeriksaan otomatis setelah restart sebelum mengirim ulang."
                elif ulang:
                    note = tahap + " | Akses/server belum siap; restart otomatis dan ulang baris ini."
            finally:
                for tab in browser.pages:
                    if tab != assignment and not tab.is_closed() and len(browser.pages) > 1:
                        with suppress(Error):
                            tab.close()
            sheet.cell(number, status, "RETRY" if ulang and result == "ERROR" else result)
            if note:
                cell = sheet.cell(number, catatan)
                text = str(cell.value or "")
                text += ("\n" if text else "") + note if note not in text else ""
                assert len(text) <= 32767, "Catatan melebihi batas sel Excel"
                cell.value, cell.data_type = text, "s"
            simpan(book, path)
            hasil[result] += 1
            print(f"  [HASIL] {sheet.cell(number, status).value}; progres disimpan.", flush=True)
            if note:
                print("  [CATATAN] " + note.splitlines()[0], flush=True)
            if ulang:
                with suppress(Error):
                    sesi(browser, assignment, root / ".fasih-profile" / "sesi.json", simpan=True)
                ulang_browser = True
                break
            page = assignment
            page.bring_to_front()
            if index + 1 < len(rows):
                istirahat = sisa == 1
                waktu, sisa = jeda(cfg, sisa)
                if istirahat:
                    hitung_mundur(page, waktu)
                else:
                    print(f"  Jeda {waktu / 1000:.1f} detik sebelum baris berikutnya.", flush=True)
                    page.wait_for_timeout(waktu)
        else:
            print(f"[SELESAI] SUCCESS={hasil['SUCCESS']}, NO_CHANGE={hasil['NO_CHANGE']}, ERROR={hasil['ERROR']}.", flush=True)
    browser.remove_listener("page", tab_baru)
    book.close()
    return ulang_browser

def main():
    pemulihan = False
    while True:
        try:
            if not jalankan(pemulihan=pemulihan):
                return
        except Error as error:
            print("[PULIH] Browser/koneksi gagal: " + str(error).splitlines()[0], flush=True)
        pemulihan = True
        print("[PULIH] Browser ditutup. Membuka ulang dengan profil/cookie sama setelah jeda 1–5 detik; Ctrl+C untuk berhenti.", flush=True)
        hitung_mundur(None, random.uniform(1000, 5000), "Menunggu restart")

if __name__ == "__main__":
    main()
