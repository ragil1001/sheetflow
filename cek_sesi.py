import json
from contextlib import contextmanager, nullcontext
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import call, patch

from camoufox.sync_api import Camoufox
from openpyxl import Workbook, load_workbook
from playwright.sync_api import Error, expect

from fasih import FIELDS, PESAN_BOT, buka_usaha, identitas, jalankan, kirim, main, pasang_refresh, pesan_galat, sesi


for sequence in ([False], [True, True, True, False]):
    with patch("fasih.jalankan", side_effect=sequence) as run, patch("fasih.sleep") as pause:
        main()
        assert run.call_args_list == [call(pemulihan=False)] + [call(pemulihan=True)] * (len(sequence) - 1)
        assert pause.call_args_list == [call(2)] * (len(sequence) - 1)
with patch("fasih.jalankan", side_effect=Error("NS_ERROR_UNKNOWN_HOST")), patch("fasih.sleep") as pause:
    with TestCase().assertRaises(Error):
        main()
    pause.assert_not_called()

modal = '<div class="tw:border! tw:relative tw:flex tw:w-full tw:flex-col tw:gap-4 tw:bg-background tw:p-6 tw:pt-4 tw:shadow-lg! tw:sm:rounded-lg">'
form = ('<button>Ringkasan</button><button onclick="this.hidden=true"><svg class="tabler-icon-edit" width="24" height="24"></svg></button>'
        '<div class="fasih-form-sidebar"><div title="SE2026 - L BLOK II">BLOK II</div><div title="CATATAN">CATATAN</div></div>'
        '<div id="se2026_nested"><div data-nested-view="true" onclick="this.dataset.clicked=1">USAHA CONTOH</div></div>'
        '<div id="nama_usaha_edit#1002"><input readonly value="NAMA FORM BERBEDA"></div>'
        '<label>Pekerja laki-laki<input value="0"></label>'
        '<div id="cek_anomali_button"><input role="switch" aria-checked="true"></div>')
submit = ('<script>let a=0,b=0,c=0;function gagal(stage){document.querySelector("#m").hidden=true;document.querySelector("#k").hidden=true;'
          'const first=document.querySelector("#awal");first.hidden=false;first.disabled=true;setTimeout(()=>{first.disabled=false},20);'
          'document.body.dataset[stage]=Number(document.body.dataset[stage]||0)+1}</script>'
          '<button id="awal" onclick="this.disabled=true;setTimeout(()=>{if(a++<3){gagal(\'awal\')}else{this.hidden=true;this.disabled=false;document.querySelector(\'#m\').hidden=false}},20)">Kirim</button>'
          '<section id="m" hidden>' + modal + '<button onclick="this.disabled=true;setTimeout(()=>{this.disabled=false;if(b++<2){gagal(\'modal\')}else{document.querySelector(\'#m\').hidden=true;document.querySelector(\'#k\').hidden=false}},20)">Kirim</button></div></section>'
          '<section id="k" hidden><p>Konfirmasi Kirim</p><button onclick="this.disabled=true;setTimeout(()=>{this.disabled=false;if(c++<2){gagal(\'konfirmasi\')}else{document.querySelector(\'#k\').hidden=true;document.querySelector(\'#s\').hidden=false}},20)">Konfirmasi</button></section>'
          '<section id="s" hidden><p>Assignment Berhasil DisubmitPerubahan Anda berhasil disimpan.</p><button>Tetap di Halaman</button></section>')

with TemporaryDirectory() as temp, patch("builtins.print") as output:
    root = Path(temp)
    fingerprint = identitas(root / "fingerprint.json")
    assert identitas(root / "fingerprint.json") == fingerprint
    cfg = dict(excel_file="uji.xlsx", sheet_name="Sheet1", delay_min=0, delay_max=0, batch_min=20, batch_max=20, rest_min=0, rest_max=0)
    (root / "config.json").write_text(json.dumps(cfg), encoding="utf-8")

    def workbook(rows):
        book = Workbook()
        sheet = book.active
        sheet.title = "Sheet1"
        sheet.append(["id_fasih", "target_usaha", *FIELDS, "status", "catatan"])
        for ident, nama, value in rows:
            sheet.append([ident, nama, *[value if key == "tk_laki_laki" else None for key in FIELDS], "TODO"])
        book.save(root / "uji.xlsx")
        book.close()

    with Camoufox(headless=True, os="windows", fingerprint=fingerprint, i_know_what_im_doing=True) as browser:
        context = browser.new_context()
        context.on("page", pasang_refresh)
        context.route("**/*", lambda r: r.fulfill(content_type="text/html", body="Survei"))
        page = context.new_page()
        page.goto("https://fasih-sm.bps.go.id/app/surveys?page=0")
        page.evaluate("() => {localStorage.setItem('uji','lokal');sessionStorage.setItem('uji','lokal')}")
        context.add_cookies([dict(name="uji", value="lokal", url=page.url)])
        snapshot = root / "sesi.json"
        sesi(context, page, snapshot, simpan=True)
        assert json.loads(snapshot.read_text(encoding="utf-8"))["tab"] == {"uji": "lokal"}
        for target, kartu in (("USAHA CONTOH (PEMILIK A)", "USAHA CONTOH"), ("USAHA CONTOH", "USAHA CONTOH &lt;PEMILIK B&gt;"), ("USAHA CONTOH <PEMILIK A", "USAHA CONTOH (LAIN)")):
            page.set_content(form.replace(">USAHA CONTOH</div>", ">" + kartu + "</div>"))
            buka_usaha(page, target)
            assert page.locator('[data-clicked="1"]').count() == 1
        page.set_content(form.replace(">USAHA CONTOH</div>", ">USAHA LAIN</div>"))
        with patch("fasih.TIMEOUT_CARI_USAHA", 50), TestCase().assertRaisesRegex(ValueError, "Kartu usaha tidak cocok"):
            buka_usaha(page, "USAHA CONTOH")
        page.set_content(form.replace(">USAHA CONTOH</div>", ">USAHA BOT-12345678</div>"))
        buka_usaha(page, "USAHA BOT-12345678")
        assert not page.is_closed()
        page.set_content(submit)
        with patch("fasih.TIMEOUT", 1000):
            assert kirim(page) == ""
        assert page.evaluate("({...document.body.dataset})") == {"awal": "3", "modal": "2", "konfirmasi": "2"}
        page.set_content('<button onclick="document.body.dataset.kirim=1">Kirim</button>' + modal + '<h2>Berhasil</h2><button>Tetap di Halaman</button></div>')
        with patch.object(page, "wait_for_timeout", side_effect=AssertionError("Sukses masih ditunggu")):
            assert kirim(page) == ""
        assert page.locator("body").get_attribute("data-kirim") is None, "Sudah sukses tetapi Kirim diklik lagi"
        page.set_content('<h2>Assignment Berhasil Disubmit</h2><button>Tetap di Halaman</button>')
        with patch("fasih.TIMEOUT", 50):
            assert pesan_galat(page) == "", "Modal sukses masih menunggu modal validasi"
        page.set_content('<button onclick="this.disabled=true;setTimeout(()=>{document.body.innerHTML=\'<h2>Assignment Berhasil Disubmit</h2><button>Tetap di Halaman</button>\'},100)">Kirim</button>')
        with patch("fasih.TIMEOUT", 50), patch.object(page, "wait_for_timeout", wraps=page.wait_for_timeout) as tunggu:
            assert kirim(page) == ""
            assert all(c.args[0] <= 100 for c in tunggu.call_args_list), "Sukses tertunda oleh jeda tetap"
        page.set_content(modal + '<button disabled>Kirim</button></div><script>setTimeout(()=>{document.body.innerHTML=\'<h2>Assignment Berhasil Disubmit</h2><button>Tetap di Halaman</button>\'},100)</script>')
        with patch.object(page, "reload", side_effect=AssertionError("Loading Kirim direfresh")):
            assert kirim(page) == "", "Kirim masih menunggu tombol pada modal lama"
        page.set_content('<button>Kirim</button>' + modal + '<button>Galat 1 Perlu diperbaiki</button></div>')
        with patch("fasih.pesan_galat", return_value="Data wajib diperbaiki"):
            assert kirim(page) == "Data wajib diperbaiki"
        page.set_content(submit.replace("function gagal(stage){", "function gagal(stage){document.querySelector('#error').hidden=false;") + '<p id="error" hidden>Failed presigned URL</p>')
        with patch("fasih.TIMEOUT", 1000):
            assert kirim(page) == "", "Pesan presigned lama menghalangi Kirim ulang"
        for html in ("<p>Bot Detected</p>", "<p>Anda terdeteksi sebagai bot</p>", "<p>BOT-12345678</p>"):
            assert PESAN_BOT.search(html)
            tab = context.new_page()
            tab.set_content(html)
            with TestCase().assertRaises((Error, AssertionError)):
                expect(tab.get_by_role("button", name="Ringkasan")).to_be_visible(timeout=500)
            assert tab.is_closed()

        page.set_content('<p>Loading...</p>')
        page.evaluate("setTimeout(() => document.body.innerHTML='<button>Ringkasan</button>', 300)")
        with patch("fasih.TIMEOUT", 50), patch.object(page, "reload", side_effect=AssertionError("Loading direfresh")):
            expect(page.get_by_role("button", name="Ringkasan")).to_be_visible(timeout=0)
        tab = context.new_page()
        tab.set_content('<button onclick="document.body.dataset.refresh=Number(document.body.dataset.refresh||0)+1;if(document.body.dataset.refresh===\'3\')document.body.innerHTML=\'<button>Ringkasan</button>\'">Refresh Page</button>')
        with patch.object(tab, "reload", side_effect=AssertionError("Refresh browser dipakai")):
            expect(tab.get_by_role("button", name="Ringkasan")).to_be_visible(timeout=2000)
        assert not tab.is_closed() and tab.locator("body").get_attribute("data-refresh") == "3"
        tab.close()

        context.route("**/app/assignment-detail/*", lambda r: r.fulfill(content_type="text/html", body='<button onclick="window.open(\'/form/' + r.request.url.rsplit('/', 1)[-1] + '\')">Review</button>'))
        def halaman(route):
            ident = route.request.url.rsplit("/", 1)[-1]
            html = form + (submit if ident == "kirim" else "")
            if ident == "kirim":
                html = html.replace('<p>Assignment Berhasil DisubmitPerubahan Anda berhasil disimpan.</p>', '<h2>Assignment Berhasil Disubmit</h2><p>Perubahan Anda berhasil disimpan.</p>')
                html = html.replace('<button>Tetap di Halaman</button>', '<button onclick="localStorage.setItem(\'sukses_ditutup\',\'1\')">Tetap di Halaman</button>')
            elif ident == "lambat":
                html = '<p>Loading...</p><template>' + form + '</template><script>setTimeout(()=>{document.body.innerHTML=document.querySelector("template").innerHTML},300)</script>'
            if ident == "beda":
                html = html.replace(">USAHA CONTOH</div>", ">USAHA LAIN</div>")
            elif ident == "akses":
                html = html.replace('<button onclick="this.hidden=true"><svg class="tabler-icon-edit" width="24" height="24"></svg></button>', "")
            route.fulfill(content_type="text/html", body=html)
        context.route("**/form/*", halaman)
        workbook([("beda", "USAHA CONTOH", None), ("akses", "USAHA CONTOH", None), ("angka", "USAHA CONTOH", -1), ("lanjut", "USAHA CONTOH (PEMILIK A)", None), ("lambat", "USAHA CONTOH", None), ("kirim", "USAHA CONTOH", 1)])
        with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 1000), patch("fasih.TIMEOUT_CARI_USAHA", 50), \
             patch("fasih.Camoufox", side_effect=[nullcontext(context)]), patch("builtins.input", return_value=""), patch("fasih.sleep") as pause:
            main()
            pause.assert_not_called()
        book = load_workbook(root / "uji.xlsx")
        rows = list(book["Sheet1"].values)
        assert [row[rows[0].index("status")] for row in rows[1:]] == ["ERROR", "ERROR", "ERROR", "NO_CHANGE", "NO_CHANGE", "SUCCESS"], rows
        assert "Kartu usaha tidak cocok" in rows[1][rows[0].index("catatan")]
        assert "Akses edit tidak tersedia" in rows[2][rows[0].index("catatan")]
        assert "Angka tidak valid" in rows[3][rows[0].index("catatan")]
        assert context.pages[0].evaluate("localStorage.getItem('sukses_ditutup')") == "1"
        book.close()
        context.cookies()
        assert len(context.pages) == 1
        workbook([("diterima", "USAHA CONTOH", 1)])
        diterima = []
        def terima(route):
            diterima.append(True)
            route.fulfill(body="ok")
        context.route("**/terima", terima)
        def terkirim(route):
            if diterima:
                html = form.replace("tabler-icon-edit", "tabler-icon-rotate-clockwise").replace('input value="0"', 'input readonly value="1"')
            else:
                html = (form + '<button onclick="this.hidden=true;document.querySelector(\'#m\').hidden=false">Kirim</button>'
                        '<section id="m" hidden>' + modal + '<button onclick="document.querySelector(\'#m\').hidden=true;document.querySelector(\'#k\').hidden=false">Kirim</button></div></section>'
                        '<section id="k" hidden><p>Konfirmasi Kirim</p><button onclick="fetch(\'/terima\').then(()=>{document.body.innerHTML=\'<p>Bot Detected</p>\'})">Konfirmasi</button></section>')
            route.fulfill(content_type="text/html", body=html)
        context.route("**/form/diterima", terkirim)
        with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 1000), patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input", return_value=""):
            assert jalankan() is True
            assert not jalankan(pemulihan=True)
        assert diterima == [True], "Kiriman yang sudah diterima dikirim ulang"
        book = load_workbook(root / "uji.xlsx")
        assert book["Sheet1"].cell(2, len(FIELDS) + 3).value == "SUCCESS"
        book.close()
        context.close()

    workbook([("pulih", "USAHA CONTOH", None)])
    launches, closed = [], []
    @contextmanager
    def buka(**options):
        assert len(launches) == len(closed) and len(launches) < 3, output.call_args_list[-20:]
        with Camoufox(**{**options, "headless": True, "timeout": 90000}) as context:
            launches.append(options["user_data_dir"])
            context.route("**/*", lambda r: r.fulfill(content_type="text/html", body="Survei"))
            def survei(route):
                route.fulfill(content_type="text/html", body="Bot Detected" if len(launches) < 3 else "Survei")
            context.route("**/app/surveys?*", survei)
            context.route("**/app/assignment-detail/*", lambda r: r.fulfill(content_type="text/html", body='<button onclick="window.open(\'/form\')">Review</button>'))
            context.route("**/form", lambda r: r.fulfill(content_type="text/html", body=form))
            yield context
        closed.append(True)
        assert not context.pages
    with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 5000), patch("fasih.Camoufox", side_effect=buka), \
         patch("builtins.input", return_value=""), patch("fasih.sleep") as pause:
        main()
        assert len(launches) == len(closed) == 3 and len(set(launches)) == 1
        assert pause.call_args_list == [call(2)] * 2
    book = load_workbook(root / "uji.xlsx")
    assert book["Sheet1"].cell(2, len(FIELDS) + 3).value == "NO_CHANGE"
    book.close()

print("Pemeriksaan kartu, Kirim berulang, akses edit, dan restart browser lulus.")
