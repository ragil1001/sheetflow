import json
from contextlib import contextmanager, nullcontext, suppress
from tempfile import TemporaryDirectory
from pathlib import Path
from unittest import TestCase
from unittest.mock import call, patch
from camoufox.sync_api import Camoufox
from openpyxl import Workbook, load_workbook
from playwright.sync_api import Error, expect
from fasih import FIELDS, PESAN_BOT, SELECTORS, buka_usaha, identitas, inputan, isi_remarks, jalankan, main, pasang_refresh, pesan_galat, sesi, tab_survei, tunggu_kirim

pesan_bot = ("BOT-12345678", "Bot Detected", "Anda terdeteksi sebagai bot", "Terdeteksi bot",
             "Kami mendeteksi perilaku yang tidak wajar pada koneksi Anda", "Sistem kami mendeteksi koneksi Anda sebagai bot")
assert all(PESAN_BOT.search(text) for text in pesan_bot)
assert not PESAN_BOT.search("USAHA ROBOTIKA")
for sequence in ([None], [True, True, True, None], [Error("Server down"), True, None]):
    with patch("fasih.jalankan", side_effect=sequence) as run, patch("fasih.hitung_mundur") as delay, patch("builtins.print"):
        main()
        assert run.call_count == len(sequence), "Pemulihan berhenti sebelum server pulih"
        assert delay.call_count == len(sequence) - 1
        assert all(1000 <= c.args[1] <= 5000 for c in delay.call_args_list)
        assert run.call_args_list == [call(pemulihan=False)] + [call(pemulihan=True)] * (len(sequence) - 1)

def nama_form(target):
    return ('<div id="nama_usaha_edit#1002"><div class="tw:flex tw:gap-2 tw:rounded-lg tw:py-2 tw:@lg/form:flex-row tw:flex-col">'
            + '<div>Nama usaha/perusahaan</div><input readonly value="' + target + '"></div></div>')

with TemporaryDirectory() as temp, patch("fasih.hitung_mundur", side_effect=lambda page, waktu, label="Istirahat batch": page.wait_for_timeout(waktu)):
    path = Path(temp) / "sesi.json"
    for run in range(2):
        fingerprint = identitas(Path(temp) / "fingerprint.json")
        assert identitas(Path(temp) / "fingerprint.json") == fingerprint
        with Camoufox(headless=True, os="windows", fingerprint=fingerprint, i_know_what_im_doing=True) as browser:
            context = browser.new_context()
            context.on("page", pasang_refresh)
            context.route("**/*", lambda r: r.fulfill(body="<!doctype html><title>Uji lokal</title><body>Uji lokal</body>"))
            page = context.new_page()
            sesi(context, page, path)
            page.goto("https://fasih-sm.bps.go.id")
            identity = page.evaluate("() => [navigator.userAgent, navigator.platform, navigator.hardwareConcurrency, screen.width, screen.height, screen.colorDepth, window.screenX, window.screenY]")
            if run == 0:
                original = identity
            else:
                assert identity == original, "Identitas berubah antar-peluncuran"
            if run == 0:
                lama = page
                baru = context.new_page()
                baru.goto("https://fasih-sm.bps.go.id/app/surveys?page=0&perPage=10&layout=list")
                lama.close()
                page = tab_survei(context)
                assert page == baru and not page.is_closed()
                page.goto("https://fasih-sm.bps.go.id/app/assignment-detail/contoh")
                context.add_cookies([{"name": "contoh", "value": "lokal", "url": page.url}])
                page.evaluate("() => { localStorage.setItem('contoh','lokal'); sessionStorage.setItem('contoh','lokal'); }")
                sesi(context, page, path, simpan=True)
            else:
                assert any(c["name"] == "contoh" and c["value"] == "lokal" for c in context.cookies())
                assert page.evaluate("localStorage.getItem('contoh')") == "lokal"
                assert page.evaluate("sessionStorage.getItem('contoh')") == "lokal"
                page.evaluate("sessionStorage.setItem('contoh','baru')")
                page.reload()
                assert page.evaluate("sessionStorage.getItem('contoh')") == "baru"
            assignment = page
            assignment.set_content('<button onclick="window.open(\'/review\')">Review</button>')
            with assignment.expect_popup() as popup:
                assignment.get_by_role("button", name="Review", exact=True).click()
            page = popup.value
            page.set_content('<button type="button">Ringkasan</button>'
                             '<button data-state="closed" onclick="this.dataset.clicked=\'true\'">'
                             '<svg class="tabler-icon tabler-icon-edit f:size-5" width="24" height="24"></svg></button>')
            page.get_by_role("button", name="Ringkasan", exact=True).wait_for()
            edit = page.locator(SELECTORS["edit"])
            edit.click()
            assert edit.get_attribute("data-clicked") == "true"
            if run == 0:
                for html in ('<button>Refresh Page</button>', '<button disabled>Refresh Page</button>', '<div id="se2026_nested"><p>Bot Detected</p></div>',
                             '<p>Gagal Mendapatkan Akses Approve/Reject/Edit Assignment</p>', *('<p>' + text + '</p>' for text in pesan_bot)):
                    tab = context.new_page()
                    tab.set_content(html)
                    with TestCase().assertRaises((Error, AssertionError)):
                        expect(tab.get_by_role("button", name="Ringkasan", exact=True)).to_be_visible(timeout=1000)
                    assert tab.is_closed(), "Halaman gangguan tidak ditutup"
                kartu_usaha = '<div data-nested-view="true" onclick="this.dataset.clicked=\'true\'; document.querySelector(\'#nama\').hidden=false"><span>USAHA CONTOH </span></div>'
                wadah_nama = '<div class="tw:flex tw:gap-2 tw:rounded-lg tw:py-2 tw:@lg/form:flex-row tw:flex-col"><div>Nama usaha/perusahaan</div>'
                for target, nama, suffix in (("USAHA CONTOH <PEMILIK A>", '<input readonly value="USAHA CONTOH (PEMILIK A)">', "1002"),
                                              ("USAHA CONTOH (PEMILIK A)", '<div>USAHA CONTOH &lt;PEMILIK A&gt;</div>', "2099")):
                    page.set_content('<div id="se2026_nested">' + kartu_usaha + '<div data-nested-view="true">USAHA LAIN</div></div>'
                                     f'<section id="nama" hidden><div id="nama_usaha_edit#{suffix}">' + wadah_nama + nama + '</div></div></section>')
                    buka_usaha(page, target)
                    assert page.locator('[data-clicked="true"]').count() == 1
                page.set_content('<div id="se2026_nested">' + kartu_usaha.replace("USAHA CONTOH", "USAHA CONTOH DUA")
                                 + '</div><section id="nama" hidden><div id="nama_usaha_edit#1002">' + wadah_nama
                                 + '<div>USAHA CONTOH DUA (PEMILIK CONTOH DUA)</div></div></div></section>')
                buka_usaha(page, "USAHA CONTOH DUA <PEMILIK CONTOH DUA")
                assert page.locator('[data-clicked="true"]').count() == 1
                page.set_content('<div id="se2026_nested">' + kartu_usaha * 2 + '</div>')
                with TestCase().assertRaisesRegex(AssertionError, "beberapa kartu"):
                    buka_usaha(page, "USAHA CONTOH <PEMILIK A>")
                assert page.locator('[data-clicked="true"]').count() == 0
                page.set_content('<div id="se2026_nested">' + kartu_usaha + '</div><section id="nama" hidden>'
                                 '<div id="nama_usaha_edit#1002">' + wadah_nama + '<input value="USAHA CONTOH (PEMILIK B)"></div></div></section>')
                with TestCase().assertRaisesRegex(ValueError, "tidak sesuai"):
                    buka_usaha(page, "USAHA CONTOH <PEMILIK A>")
                page.set_content('<div id="se2026_nested">' + kartu_usaha + '</div><section id="nama" hidden>'
                                 '<div id="nama_usaha_edit#1002">' + wadah_nama + '<div>USAHA CONTOH (PEMILIK B)</div></div></div></section>'
                                 '<div>USAHA CONTOH &lt;PEMILIK A&gt;</div>')
                with TestCase().assertRaisesRegex(ValueError, "tidak sesuai"):
                    buka_usaha(page, "USAHA CONTOH <PEMILIK A>")
                page.set_content('<div id="se2026_nested">' + kartu_usaha
                                 + '<div data-nested-view="true" onclick="this.dataset.clicked=\'true\'"><span>USAHA CONTOH &lt;PEMILIK A&gt;</span></div></div>' + nama_form("USAHA CONTOH (PEMILIK A)"))
                buka_usaha(page, "USAHA CONTOH <PEMILIK A>")
                assert page.locator('[data-clicked="true"]').inner_text() == "USAHA CONTOH <PEMILIK A>"
                page.set_content('<div id="se2026_nested"><div data-nested-view="true" onclick="document.getElementById(\'nama_usaha_edit#1002\').hidden=false">USAHA SIAP</div></div>'
                                 '<div id="nama_usaha_edit#1002" hidden><label>Nama usaha/perusahaan<input value="USAHA SIAP"></label></div>')
                buka_usaha(page, "USAHA SIAP")
                page.set_content('<div id="se2026_nested"><div data-nested-view="true">USAHA SIAP</div></div><input id="nama_usaha_edit#1002" value="USAHA SIAP">')
                buka_usaha(page, "USAHA SIAP")
                for form in ("", '<input id="nama_usaha_edit#1002" value="">', nama_form("USAHA SIAP") * 2):
                    page.set_content('<div id="se2026_nested"><div data-nested-view="true">USAHA SIAP</div></div>' + form)
                    with patch("fasih.TIMEOUT", 100), TestCase().assertRaisesRegex(ValueError, "Nama usaha dalam form tidak dapat diverifikasi"):
                        buka_usaha(page, "USAHA SIAP")
                page.set_content('<div id="se2026_nested"><div data-nested-view="true">USAHA BOT-12345678</div></div>' + nama_form("USAHA BOT-12345678"))
                buka_usaha(page, "USAHA BOT-12345678")
                assert not page.is_closed(), "Teks kartu/nama usaha dianggap pesan BOT"
                page.set_content('<div id="se2026_nested"><div data-nested-view="true"><div data-nested-view="true" '
                                 'onclick="this.dataset.clicked=\'true\'"><span>USAHA&nbsp;CONTOH  TIGA</span></div></div></div>' + nama_form("USAHA CONTOH TIGA"))
                buka_usaha(page, "USAHA CONTOH TIGA")
                assert page.locator('[data-clicked="true"]').count() == 1
                page.set_content('<div id="se2026_nested"><div data-nested-view="true" onclick="this.dataset.clicked=\'true\'">USAHA KEMBAR</div>'
                                 '<div data-nested-view="true" onclick="this.dataset.clicked=\'true\'">USAHA KEMBAR</div></div>')
                with TestCase().assertRaisesRegex(AssertionError, "beberapa kartu"):
                    buka_usaha(page, "USAHA KEMBAR")
                assert page.locator('[data-clicked="true"]').count() == 0
                page.set_content('<div id="se2026_nested"><div data-nested-view="true" onclick="this.dataset.clicked=\'true\'">USAHA LAIN</div></div>')
                with patch("fasih.TIMEOUT", 2000), patch("fasih.TIMEOUT_CARI_USAHA", 100):
                    with TestCase().assertRaisesRegex(ValueError, "Kartu usaha tidak ditemukan: USAHA TARGET") as gagal:
                        buka_usaha(page, "USAHA TARGET")
                    assert "Kartu terlihat: USAHA LAIN" in str(gagal.exception)
                assert page.locator('[data-clicked="true"]').count() == 0
                page.set_content('<div id="se2026_nested"></div>' + nama_form("USAHA LAMBAT"))
                page.evaluate("html => setTimeout(() => document.querySelector('#se2026_nested').innerHTML=html, 250)",
                              '<div data-nested-view="true" onclick="this.dataset.clicked=\'true\'">USAHA LAMBAT</div>')
                with patch("fasih.TIMEOUT", 2000), patch("fasih.TIMEOUT_CARI_USAHA", 100):
                    buka_usaha(page, "USAHA LAMBAT")
                assert page.locator('[data-clicked="true"]').inner_text() == "USAHA LAMBAT"
                page.set_content('<div id="se2026_nested"><div data-nested-view="true">USAHA LAIN</div></div>' + nama_form("USAHA SUSULAN (PEMILIK A)"))
                page.evaluate("html => setTimeout(() => document.querySelector('#se2026_nested').insertAdjacentHTML('beforeend',html), 250)",
                              '<div data-nested-view="true" onclick="this.dataset.clicked=\'true\'">USAHA SUSULAN (PEMILIK A)</div>')
                with patch("fasih.TIMEOUT_CARI_USAHA", 1000):
                    buka_usaha(page, "USAHA SUSULAN <PEMILIK A>")
                assert page.locator('[data-clicked="true"]').inner_text() == "USAHA SUSULAN (PEMILIK A)"
                page.set_content('<div id="se2026_nested"></div>')
                with patch("fasih.TIMEOUT", 100), patch("fasih.TIMEOUT_CARI_USAHA", 10):
                    with TestCase().assertRaisesRegex(AssertionError, "Daftar kartu usaha BLOK II tidak muncul") as gagal:
                        buka_usaha(page, "USAHA TARGET")
                    assert "timeout 100ms" in str(gagal.exception)
                button_note = '<button class="tw:size-8 tw:relative tw:text-primary tw:@lg/form:flex! tw:hidden" onclick="document.querySelector(\'#editor\').hidden=false; document.body.dataset.field=this.parentElement.id"><svg class="tw:size-4" width="16" height="16"></svg></button>'
                page.set_content('<div id="gaji#2099"><div class="tw:rounded-lg tw:gap-2"><div>26. a. Total upah dan gaji, serta jaminan sosial pegawai*Rp</div>'
                                 '<input value="0"></div>' + button_note
                                 + '<button class="tw:size-8 tw:relative tw:text-primary" onclick="document.body.dataset.salah=\'true\'"><svg class="tw:size-4" width="16" height="16"></svg></button>'
                                 + '</div><div id="biaya#1002">' + button_note + '</div>'
                                 '<section id="editor" hidden><textarea placeholder="Tambah catatan di sini..."></textarea>'
                                 '<button onclick="document.body.dataset.remark=document.querySelector(\'textarea\').value">Simpan<svg></svg></button>'
                                 '<button onclick="this.parentElement.hidden=true">Dismiss<svg></svg></button></section>')
                isi_remarks(page, inputan(page, FIELDS["upah_gaji"]), "Perbaikan hasil pemeriksaan\nNilai sesuai konfirmasi")
                assert page.locator("body").get_attribute("data-field") == "gaji#2099"
                assert page.locator("body").get_attribute("data-salah") is None
                assert page.locator("body").get_attribute("data-remark") == "Perbaikan hasil pemeriksaan\nNilai sesuai konfirmasi"
                modal = '<div class="tw:border! tw:relative tw:flex tw:w-full tw:flex-col tw:gap-4 tw:bg-background tw:p-6 tw:pt-4 tw:shadow-lg! tw:sm:rounded-lg">'
                kartu = ('<div class="tw:flex tw:items-start tw:justify-between tw:gap-3 tw:rounded-lg tw:border tw:border-border/50 tw:bg-card tw:p-3 tw:shadow-3xs tw:transition-all tw:hover:border-border/80 tw:hover:shadow-2xs">'
                         '<div class="tw:flex tw:min-w-0 tw:flex-1 tw:items-start tw:gap-2">'
                         '<div>1</div><div><div>SE2026 - L BLOK I</div>'
                         '<div title="9. b. Alamat domisili:">9. b. Alamat domisili:</div>'
                         '<ul class="tw:mt-1.5 tw:list-outside tw:space-y-0.5 tw:pl-2.5 tw:text-[11px] tw:text-destructive">'
                         '<li class="tw:list-disc">Jika umur kurang dari 17 tahun, maka alamat domisili harus berkode 1 atau 4</li>'
                         '</ul></div></div><button title="Lihat">Lihat</button></div>')
                pesan = "Jika umur kurang dari 17 tahun, maka alamat domisili harus berkode 1 atau 4"
                for jumlah in (1, 2):
                    page.set_content(modal + '<button onclick="this.parentElement.hidden=true; document.querySelector(\'#rincian\').hidden=false">'
                                     f'Galat {jumlah} Perlu diperbaiki</button><button disabled>Kirim</button></div>'
                                     '<section id="rincian" hidden>' + modal + f'<div>Total: {jumlah} / {jumlah}</div>'
                                     + kartu * jumlah + '</div></section><ul class="tw:text-destructive"><li>Bukan pesan modal</li></ul>')
                    assert pesan_galat(page) == "\n".join(["9. b. Alamat domisili: — " + pesan] * jumlah)
                    assert page.get_by_role("button", name="Kirim", exact=True).is_hidden()
                page.set_content(modal + '<button onclick="this.parentElement.hidden=true; document.querySelector(\'#rincian\').hidden=false">Galat 1 Perlu diperbaiki</button></div>'
                                 '<section id="rincian" hidden>' + modal + '<div>Total: 1 / 1</div>'
                                 + kartu.replace('<div title="9. b. Alamat domisili:">9. b. Alamat domisili:</div>', '<div>Pilih UMKM dalam satu SLS yang sama</div>')
                                 .replace(pesan, "Wajib diisi") + '</div></section>')
                assert "Pilih UMKM dalam satu SLS yang sama — Wajib diisi" in pesan_galat(page)
                for isi in ('<button>Galat 0 Perlu diperbaiki</button>', ''):
                    page.set_content(modal + isi + '<button onclick="this.dataset.clicked=\'true\'">Kirim</button></div>')
                    assert pesan_galat(page) == ""
                    kirim = page.get_by_role("button", name="Kirim", exact=True)
                    assert kirim.get_attribute("data-clicked") is None
                    kirim.click()
                    assert kirim.get_attribute("data-clicked") == "true"
                page.set_content('<p hidden>Failed presigned URL</p><p>Konfirmasi Kirim</p>')
                tunggu_kirim(page, page.get_by_text("Konfirmasi Kirim", exact=True))
                for text in ("Failed presigned URL", "Gagal mendapatkan pre-signed URL", "Presigned URL request failed"):
                    page.set_content('<button onclick="this.dataset.clicked=\'true\'">Kirim</button><p hidden></p>')
                    page.evaluate("text => setTimeout(() => { const p=document.querySelector('p'); p.textContent=text; p.hidden=false; }, 50)", text)
                    with patch("fasih.TIMEOUT", 1000), TestCase().assertRaisesRegex(AssertionError, "kesalahan presigned URL"):
                        tunggu_kirim(page, page.get_by_text("Konfirmasi Kirim", exact=True))
                    assert page.get_by_role("button", name="Kirim", exact=True).get_attribute("data-clicked") is None
                page.set_content('<p>Failed presigned URL</p>')
                with TestCase().assertRaisesRegex(AssertionError, "presigned URL"):
                    pesan_galat(page)
            page.close()
            page = assignment
            page.bring_to_front()
            assert context.pages == [assignment] and not assignment.is_closed()
            if run == 0:
                root = Path(temp) / "uji-error"
                root.mkdir()
                cfg = dict(excel_file="uji.xlsx", sheet_name="Sheet1", delay_min=0, delay_max=0,
                           batch_min=2, batch_max=2, rest_min=0, rest_max=0, action_min=0, action_max=0, typing_min=1, typing_max=1)
                (root / "config.json").write_text(json.dumps(cfg), encoding="utf-8")
                book = Workbook()
                sheet = book.active
                sheet.title = "Sheet1"
                sheet.append(["id_fasih", "target_usaha", *FIELDS, "remark", "status", "catatan"])
                sheet.append(["macet", "USAHA CONTOH TIGA", *([None] * len(FIELDS)), None, "TODO", "Catatan operator"])
                sheet.append(["lanjut", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), None, "TODO"])
                sheet.append(["ubah", "USAHA BERIKUTNYA", *[0 if k in ("tk_laki_laki", "upah_gaji") else None for k in FIELDS], "Hasil konfirmasi", "TODO"])
                sheet.append(["tetap", "USAHA BERIKUTNYA", *[0 if k in ("tk_laki_laki", "upah_gaji") else None for k in FIELDS], "Tidak disalin", "TODO"])
                sheet.append(["input-hilang", "USAHA BERIKUTNYA", *[1 if k == "perbaikan_produksi" else None for k in FIELDS], None, "TODO"])
                sheet.append(["nama-berbeda", "USAHA timeout (PEMILIK A)", *([None] * len(FIELDS)), None, "TODO"])
                sheet.append(["nama-hilang", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), None, "TODO"])
                sheet.append(["nama-kosong", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), None, "TODO"])
                sheet.append(["nama-bot-data", "USAHA BOT-12345678", *([None] * len(FIELDS)), None, "TODO"])
                sheet.append(["sesudah-error", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), None, "TODO"])
                book.save(root / "uji.xlsx")
                book.close()
                context.route("**/app/assignment-detail/*", lambda r: r.fulfill(content_type="text/html",
                              body='<button onclick="window.open(\'/uji-form/' + r.request.url.rsplit("/", 1)[-1] + '\')">Review</button>'))
                context.route("**/uji-form/*", lambda r: r.fulfill(content_type="text/html", body=(
                    '<button>Ringkasan</button><button onclick="this.hidden=true">'
                    '<svg class="tabler-icon tabler-icon-edit" width="24" height="24"></svg></button>'
                    '<div class="fasih-form-sidebar"><div title="SE2026 - L BLOK II">BLOK II</div>'
                    '<div title="CATATAN">CATATAN</div></div><div id="se2026_nested">'
                    '<div data-nested-view="true"><span>USAHA BERIKUTNYA</span></div></div>'
                    + nama_form("USAHA BERIKUTNYA") + '<div id="cek_anomali_button"><input role="switch" aria-checked="true"></div>'
                    + ''.join('<div id="' + key + '#2099"><div class="tw:rounded-lg tw:gap-2"><div>' + label + '</div><input value="'
                              + ('1' if r.request.url.endswith('/ubah') else '0') + '"></div>' + button_note + '</div>'
                              for key, label in (("pekerja", FIELDS["tk_laki_laki"]), ("gaji", FIELDS["upah_gaji"])))
                    + '<section id="editor" hidden><textarea placeholder="Tambah catatan di sini..."></textarea>'
                    '<button onclick="localStorage.setItem(\'remarks_\'+document.body.dataset.field,document.querySelector(\'textarea\').value)">Simpan</button>'
                    '<button onclick="this.parentElement.hidden=true">Dismiss</button></section><button>Kirim</button>')))
                def nama_bermasalah(route):
                    kasus = route.request.url.rsplit("/", 1)[-1]
                    kartu = "USAHA timeout" if kasus == "nama-berbeda" else "USAHA BOT-12345678" if kasus == "nama-bot-data" else "USAHA BERIKUTNYA"
                    form = "" if kasus == "nama-hilang" else nama_form("" if kasus == "nama-kosong" else "USAHA timeout (PEMILIK B)")
                    route.fulfill(content_type="text/html", body=(
                        '<button>Ringkasan</button><button onclick="this.hidden=true"><svg class="tabler-icon-edit" width="24" height="24"></svg></button>'
                        '<div class="fasih-form-sidebar"><div title="SE2026 - L BLOK II">BLOK II</div></div>'
                        '<div id="se2026_nested"><div data-nested-view="true">' + kartu + '</div></div>' + form))
                context.route("**/uji-form/nama-*", nama_bermasalah)
                with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 5000), patch("fasih.TIMEOUT_CARI_USAHA", 100), \
                     patch("fasih.Camoufox", side_effect=[nullcontext(context)]) as launch, patch("builtins.input", return_value=""), \
                     patch("fasih.isi_remarks", wraps=isi_remarks) as remarks, patch("fasih.pesan_galat", return_value="Validasi uji") as validasi, \
                     patch("builtins.print") as output:
                    main()
                    assert remarks.call_count == 2
                    assert all(c.args[2] == "Hasil konfirmasi" for c in remarks.call_args_list)
                    assert validasi.call_count == 1, "Input hilang diteruskan ke Kirim"
                    pesan_terminal = [c.args[0] for c in output.call_args_list]
                    assert any("[MULAI] 10 baris; headless=" in text for text in pesan_terminal)
                    assert "[SESI] Halaman Survei terbuka. Mulai memeriksa assignment." in pesan_terminal
                    assert "[1/10] Baris Excel 2: USAHA CONTOH TIGA" in pesan_terminal
                    assert pesan_terminal.count("  [HASIL] NO_CHANGE; progres disimpan.") == 3
                    assert pesan_terminal.count("  [HASIL] ERROR; progres disimpan.") == 7
                    assert "  [CATATAN] Galat validasi Kirim: Validasi uji" in pesan_terminal
                    assert pesan_terminal[-1] == "[SELESAI] SUCCESS=0, NO_CHANGE=3, ERROR=7."
                    assert not any("[PULIH]" in text for text in pesan_terminal), "Nama usaha bermasalah memicu restart"
                    launch.assert_called_once()
                book = load_workbook(root / "uji.xlsx")
                data = list(book["Sheet1"].values)
                headers = list(data[0])
                assert data[1][headers.index("status")] == "ERROR"
                note = data[1][headers.index("catatan")]
                assert note.startswith("Catatan operator\nBLOK II: membuka usaha USAHA CONTOH TIGA | ValueError:")
                assert "Kartu usaha tidak ditemukan" in note
                assert "Kartu terlihat: USAHA BERIKUTNYA" in note
                assert data[2][headers.index("status")] == "NO_CHANGE"
                assert data[3][headers.index("status")] == "ERROR"
                assert data[3][headers.index("keterangan")] == "Validasi uji"
                assert data[4][headers.index("status")] == "NO_CHANGE"
                assert data[5][headers.index("status")] == "ERROR"
                assert "Input tidak terlihat/unik setelah nama usaha muncul" in data[5][headers.index("catatan")]
                assert "timeout" not in data[5][headers.index("catatan")].lower(), "Input hilang masih ditunggu"
                assert data[6][headers.index("status")] == "ERROR" and "Nama usaha dalam form tidak sesuai" in data[6][headers.index("catatan")]
                assert "Nama form: USAHA timeout (PEMILIK B)" in data[6][headers.index("catatan")]
                for row in data[7:9]:
                    assert row[headers.index("status")] == "ERROR" and "Nama usaha dalam form tidak dapat diverifikasi" in row[headers.index("catatan")]
                assert data[9][headers.index("status")] == "ERROR" and "Nama usaha dalam form tidak sesuai" in data[9][headers.index("catatan")]
                assert data[10][headers.index("status")] == "NO_CHANGE", "Error input/nama menghentikan baris selanjutnya"
                assignment.wait_for_function("() => ['gaji','pekerja'].every(k => localStorage.getItem('remarks_'+k+'#2099') === 'Hasil konfirmasi')", timeout=5000)
                context.cookies()
                assert context.pages == [assignment], [tab.url for tab in context.pages]
                book.close()
                context.unroute("**/uji-form/nama-*", nama_bermasalah)
                for panne in ("bot", "down", "popup-bot", "popup-down", "refresh", "popup-refresh", "blok-bot", "nama-bot"):
                    book = Workbook()
                    sheet = book.active
                    sheet.title = "Sheet1"
                    sheet.append(["id_fasih", "target_usaha", *FIELDS, "status"])
                    sheet.append([panne, "USAHA BERIKUTNYA", *([None] * len(FIELDS)), "TODO"])
                    sheet.append(["suivant", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), "TODO"])
                    book.save(root / "uji.xlsx")
                    book.close()
                    if panne.startswith(("popup-", "blok-", "nama-")):
                        route = "**/uji-form/" + panne
                    else:
                        route = "**/app/assignment-detail/" + panne
                    def gangguan(route):
                        html = "<p>Bot Detected</p>" if panne.endswith("bot") else "<button>Refresh Page</button>" if panne.endswith("refresh") else "<p>Service unavailable</p>"
                        if panne.startswith(("blok-", "nama-")):
                            html = ('<button>Ringkasan</button><button onclick="this.hidden=true"><svg class="tabler-icon-edit" width="24" height="24"></svg></button>'
                                    '<div class="fasih-form-sidebar"><div title="SE2026 - L BLOK II" onclick="document.body.innerHTML=\'<p>Bot Detected</p>\'">BLOK II</div></div>')
                            if panne == "nama-bot":
                                html = html.replace(' onclick="document.body.innerHTML=\'<p>Bot Detected</p>\'"', '') + '<div id="se2026_nested"><div data-nested-view="true" onclick="document.body.innerHTML=\'<p>Bot Detected</p>\'">USAHA BERIKUTNYA</div></div>'
                        route.fulfill(content_type="text/html", body=html)
                    context.route(route, gangguan)
                    for attempt in range(3):
                        with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 500), \
                             patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input", return_value=""), patch("builtins.print") as output:
                            result = jalankan(pemulihan=attempt > 0)
                            assert result is True, (result, output.call_args_list)
                        assert len(context.pages) <= 1, "Tab gagal tertinggal"
                    book = load_workbook(root / "uji.xlsx")
                    assert book["Sheet1"].cell(2, len(FIELDS) + 3).value == "RETRY"
                    assert book["Sheet1"].cell(3, len(FIELDS) + 3).value == "TODO", "Server gagal tetapi baris berikut ikut diproses"
                    note = book["Sheet1"].cell(2, len(FIELDS) + 5).value
                    assert len(note.splitlines()) == len(set(note.splitlines())), "Catatan retry yang sama menumpuk"
                    book.close()
                    assert len(context.pages) <= 1, "Tab gagal tertinggal"
                    context.unroute(route)
                    with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 5000), \
                         patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input") as prompt, patch("builtins.print"):
                        assert not jalankan(pemulihan=True)
                        prompt.assert_not_called()
                    book = load_workbook(root / "uji.xlsx")
                    assert book["Sheet1"].cell(2, len(FIELDS) + 3).value == "NO_CHANGE", "Assignment gagal tidak dicoba ulang"
                    book.close()
                for jawaban in ("timeout", "bot", "presigned", "diterima-timeout", "diterima-bot", "sukses"):
                    book = Workbook()
                    sheet = book.active
                    sheet.title = "Sheet1"
                    sheet.append(["id_fasih", "target_usaha", *FIELDS, "status", "catatan"])
                    sheet.append(["submit", "USAHA BERIKUTNYA", *[1 if k == "tk_laki_laki" else None for k in FIELDS], "TODO"])
                    sheet.append(["lanjut-submit", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), "TODO"])
                    book.save(root / "uji.xlsx")
                    book.close()
                    kiriman = []
                    def terima_kirim(route):
                        kiriman.append(True)
                        route.fulfill(body="ok")
                    context.route("**/uji-kirim", terima_kirim)
                    html_form = ('<button>Ringkasan</button><button onclick="this.hidden=true"><svg class="tabler-icon-edit" width="24" height="24"></svg></button>'
                                 '<div class="fasih-form-sidebar"><div title="SE2026 - L BLOK II">BLOK II</div><div title="CATATAN">CATATAN</div></div>'
                                 '<div id="se2026_nested"><div data-nested-view="true">USAHA BERIKUTNYA</div></div>' + nama_form("USAHA BERIKUTNYA")
                                 + '<label>Pekerja laki-laki<input value="0"></label><div id="cek_anomali_button"><input role="switch" aria-checked="true"></div>'
                                 '<button onclick="this.hidden=true;document.querySelector(\'#modal\').hidden=false">Kirim</button>'
                                 '<section id="modal" hidden>' + modal + '<button onclick="document.querySelector(\'#konfirmasi\').hidden=false">Kirim</button></div></section>'
                                 '<section id="konfirmasi" hidden><p>Konfirmasi Kirim</p><button onclick="localStorage.setItem(\'jumlah_kirim\',String(Number(localStorage.getItem(\'jumlah_kirim\')||0)+1));fetch(\'/uji-kirim\').then(() => {'
                                 + ("document.querySelector('#selesai').hidden=false" if jawaban == "sukses" else "document.body.innerHTML='<p>Bot Detected</p>'" if jawaban.endswith("bot") else "document.body.innerHTML='<p>Failed presigned URL</p>'" if jawaban == "presigned" else "this.disabled=true") + '})'
                                 + '">Konfirmasi</button></section><section id="selesai" hidden><p>Assignment Berhasil DisubmitPerubahan Anda berhasil disimpan.</p><button>Tetap di Halaman</button></section>')
                    def hasil_kirim(route):
                        html = html_form
                        if kiriman and jawaban.startswith("diterima-"):
                            html = html.replace('<button onclick="this.hidden=true"><svg class="tabler-icon-edit" width="24" height="24"></svg></button>', '<button onclick="document.body.dataset.revoke=\'true\'"><svg class="tabler-icon-rotate-clockwise" width="24" height="24"></svg></button>').replace('<label>Pekerja laki-laki<input value="0">', '<label>Pekerja laki-laki<input readonly value="1">')
                        elif kiriman and (jawaban not in ("timeout", "bot") or len(kiriman) >= 2):
                            html = html_form.replace("document.body.innerHTML='<p>Bot Detected</p>'" if jawaban == "bot" else "document.body.innerHTML='<p>Failed presigned URL</p>'" if jawaban == "presigned" else "this.disabled=true", "document.querySelector('#selesai').hidden=false")
                            html = html.replace('<label>Pekerja laki-laki<input value="0">', '<label>Pekerja laki-laki<input value="1">')
                        route.fulfill(content_type="text/html", body=html)
                    context.route("**/uji-form/submit", hasil_kirim)
                    assignment = context.pages[0]
                    assignment.evaluate("localStorage.removeItem('jumlah_kirim')")
                    sesi(context, assignment, root / ".fasih-profile" / "sesi.json", simpan=True)
                    with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 1000), \
                         patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input", return_value=""), patch("builtins.print") as output:
                        assert jalankan() is (jawaban != "sukses")
                    book = load_workbook(root / "uji.xlsx")
                    assert book["Sheet1"].cell(2, len(FIELDS) + 3).value == ("SUCCESS" if jawaban == "sukses" else "RETRY"), (jawaban, output.call_args_list)
                    book.close()
                    for attempt in range(3):
                        with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 1000), \
                             patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.print"):
                            if not jalankan(pemulihan=True):
                                break
                    else:
                        raise AssertionError("Kirim tidak selesai otomatis: " + jawaban)
                    assert len(kiriman) == (3 if jawaban in ("timeout", "bot") else 2 if jawaban == "presigned" else 1), (jawaban, kiriman)
                    book = load_workbook(root / "uji.xlsx")
                    assert book["Sheet1"].cell(2, len(FIELDS) + 3).value == "SUCCESS", jawaban
                    assert book["Sheet1"].cell(3, len(FIELDS) + 3).value == "NO_CHANGE"
                    book.close()
                    context.unroute("**/uji-form/submit")
                    context.unroute("**/uji-kirim", terima_kirim)
                    assignment = context.pages[0]
                def tab_berlebih(tab):
                    for _ in range(4):
                        context.new_page()
                    context.cookies()
                    assert len(context.pages) <= 2, "Tab gagal tidak dibatasi"
                    return "Validasi uji"
                book = load_workbook(root / "uji.xlsx")
                book["Sheet1"].cell(2, len(FIELDS) + 3, "TODO")
                book.save(root / "uji.xlsx")
                book.close()
                with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 5000), \
                     patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input", return_value=""), \
                     patch("fasih.pesan_galat", side_effect=tab_berlebih) as validasi, patch("builtins.print"):
                    assert not jalankan()
                    validasi.assert_called_once()
                assert len(context.pages) == 1, "Tab tambahan tidak dibersihkan"
                for lokasi in ("assignment-detail", "uji-form"):
                    book = Workbook()
                    sheet = book.active
                    sheet.title = "Sheet1"
                    sheet.append(["id_fasih", "target_usaha", *FIELDS, "status", "catatan"])
                    sheet.append(["expired", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), "TODO", "Catatan operator"])
                    sheet.append(["jangan-dibuka", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), "TODO"])
                    book.save(root / "uji.xlsx")
                    book.close()
                    pola = "**/" + ("app/" if lokasi == "assignment-detail" else "") + lokasi + "/expired"
                    context.route(pola, lambda r: r.fulfill(status=302, headers={"location": "https://sso.bps.go.id/login"}))
                    permintaan = []
                    context.route("**/app/assignment-detail/jangan-dibuka", lambda r: (permintaan.append(r.request.url), r.fulfill(body="Tidak boleh dibuka")))
                    snapshot = root / ".fasih-profile" / "sesi.json"
                    tab_tersimpan = json.loads(snapshot.read_text(encoding="utf-8"))["tab"]
                    with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 1000), \
                         patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input", return_value=""), \
                         patch("builtins.print") as output:
                        assert jalankan() is True
                    state = json.loads(snapshot.read_text(encoding="utf-8"))
                    assert state["tab"] == tab_tersimpan, "SessionStorage FASIH tertimpa sessionStorage SSO"
                    assert state["browser"]["cookies"] == context.cookies(), "Cookie terkini tidak disimpan sebelum restart"
                    assert not permintaan, "Skrip lanjut ke baris berikutnya setelah logout"
                    assert any("restart otomatis" in str(c) for c in output.call_args_list)
                    assert not any("[SELESAI]" in str(c) for c in output.call_args_list)
                    book = load_workbook(root / "uji.xlsx")
                    data = list(book["Sheet1"].values)
                    headers = list(data[0])
                    assert data[1][headers.index("status")] == "RETRY"
                    assert "Catatan operator\n" in data[1][headers.index("catatan")]
                    assert "Akses/server belum siap" in data[1][headers.index("catatan")]
                    assert data[2][headers.index("status")] == "TODO" and data[2][headers.index("catatan")] is None
                    book.close()
                    for tab in context.pages:
                        if tab != assignment:
                            tab.close()
                    context.unroute(pola)
                    context.unroute("**/app/assignment-detail/jangan-dibuka")
                    with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 5000), \
                         patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input") as prompt, \
                         patch("builtins.print"):
                        assert not jalankan(pemulihan=True)
                        prompt.assert_not_called()
                    book = load_workbook(root / "uji.xlsx")
                    data = list(book["Sheet1"].values)
                    headers = list(data[0])
                    assert data[1][headers.index("status")] == "NO_CHANGE"
                    assert data[2][headers.index("status")] == "NO_CHANGE"
                    book.close()
                context.route("**/app/surveys?*", lambda r: r.fulfill(status=302, headers={"location": "https://sso.bps.go.id/login"}))
                book = load_workbook(root / "uji.xlsx")
                book["Sheet1"].cell(3, headers.index("status") + 1, "TODO")
                book.save(root / "uji.xlsx")
                book.close()
                with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 1000), \
                     patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input") as prompt, \
                     patch("fasih.simpan") as save, patch("builtins.print") as output:
                    assert jalankan(pemulihan=True) is True
                    prompt.assert_not_called()
                    save.assert_not_called()
                    assert any("browser akan dibuka ulang" in str(c) for c in output.call_args_list)
                context.route("**/app/surveys?*", lambda r: r.fulfill(status=503, body="Service unavailable"))
                with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 1000), \
                     patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input") as prompt, patch("builtins.print"):
                    assert jalankan(pemulihan=True) is True
                    prompt.assert_not_called()
                context.unroute("**/app/surveys?*")
                context.route("**/app/surveys?*", lambda r: r.fulfill(status=302, headers={"location": "https://sso.bps.go.id/login"}))
                with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 5000), \
                     patch("fasih.Camoufox", return_value=nullcontext(context)) as launch, patch("builtins.input", side_effect=EOFError("Login manual")) as prompt, \
                     patch("fasih.simpan") as simpan, patch("builtins.print") as output:
                    with suppress(EOFError):
                        main()
                simpan.assert_not_called()
                if launch.call_args.kwargs["headless"]:
                    prompt.assert_not_called()
                    assert any("Sesi FASIH berakhir" in str(c) for c in output.call_args_list)
                else:
                    prompt.assert_called_once()
            context.close()
    # Browser asli ditutup sebelum profil sementara yang sama dibuka lagi.
    launches, closed = [], []
    book = Workbook()
    sheet = book.active
    sheet.title = "Sheet1"
    sheet.append(["id_fasih", "target_usaha", *FIELDS, "status"])
    sheet.append(["restart", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), "TODO"])
    book.save(root / "uji.xlsx")
    book.close()
    (root / ".fasih-profile" / "sesi.json").unlink()
    @contextmanager
    def buka_ulang(**options):
        assert len(launches) == len(closed), "Browser lama masih terbuka saat restart"
        assert len(launches) < 4, "Pemulihan tidak selesai setelah server normal"
        with Camoufox(**{**options, "headless": True, "timeout": 90000}) as profile:
            launches.append(options["user_data_dir"])
            if len(launches) == 1:
                profile.add_cookies([{"name": "restart", "value": "lokal", "url": "https://fasih-sm.bps.go.id"}])
            profile.route("**/*", lambda r: r.fulfill(content_type="text/html", body="Survei"))
            profile.route("**/app/assignment-detail/*", lambda r: r.fulfill(content_type="text/html", body=(
                "<p>Bot Detected</p>" if len(launches) == 2 else '<button onclick="window.open(\'/form\')">Review</button>')))
            def survei(route):
                route.fulfill(content_type="text/html", status=503 if len(launches) == 3 else 200, body="<p>Bot Detected</p>" if len(launches) == 1 else "Survei")
            profile.route("**/app/surveys?*", survei)
            profile.route("**/form", lambda r: r.fulfill(content_type="text/html", body=(
                '<button>Ringkasan</button><button onclick="this.hidden=true"><svg class="tabler-icon-edit" width="24" height="24"></svg></button>'
                '<div class="fasih-form-sidebar"><div title="SE2026 - L BLOK II">BLOK II</div><div title="CATATAN">CATATAN</div></div>'
                '<div id="se2026_nested"><div data-nested-view="true">USAHA BERIKUTNYA</div></div>' + nama_form("USAHA BERIKUTNYA")
                + '<div id="cek_anomali_button"><input role="switch" aria-checked="true"></div>')))
            yield profile
        closed.append(True)
        assert not profile.pages, "Tab browser lama masih tersisa"
    def cek_restore(browser, page, path, simpan=False):
        sesi(browser, page, path, simpan=simpan)
        if not simpan:
            assert any(c["name"] == "restart" and c["value"] == "lokal" for c in browser.cookies()), "Cookie tidak dipulihkan sebelum membuka Survei"
    with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 3000), \
         patch("fasih.Camoufox", side_effect=buka_ulang), patch("fasih.sesi", side_effect=cek_restore), patch("fasih.hitung_mundur"), \
         patch("builtins.input", return_value="") as prompt, patch("builtins.print"):
        main()
        assert len(launches) == len(closed) == 4 and len(set(launches)) == 1
        prompt.assert_not_called()
    book = load_workbook(root / "uji.xlsx")
    assert book["Sheet1"].cell(2, len(FIELDS) + 3).value == "NO_CHANGE"
    book.close()
