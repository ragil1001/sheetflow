import json
from contextlib import nullcontext
from tempfile import TemporaryDirectory
from pathlib import Path
from unittest import TestCase
from unittest.mock import call, patch
from camoufox.sync_api import Camoufox
from openpyxl import Workbook, load_workbook
from playwright.sync_api import expect
from fasih import FIELDS, SELECTORS, buka_usaha, identitas, inputan, isi_remarks, main, pasang_refresh, pesan_galat, sesi, tab_survei

with TemporaryDirectory() as temp:
    path = Path(temp) / "sesi.json"
    for run in range(2):
        fingerprint = identitas(Path(temp) / "fingerprint.json")
        assert identitas(Path(temp) / "fingerprint.json") == fingerprint
        with Camoufox(headless=True, os="windows", fingerprint=fingerprint, i_know_what_im_doing=True) as browser:
            context = browser.new_context()
            context.on("page", pasang_refresh)
            context.route("**/*", lambda r: r.fulfill(body="<!doctype html><title>Uji lokal</title>"))
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
                for tab, tujuan in ((assignment, "Review"), (page, "Ringkasan")):
                    jawaban = ['<button onclick="location.reload()"><svg class="tabler-icon tabler-icon-refresh"></svg>Refresh Page</button>'] * 2 + [f'<button>{tujuan}</button>']
                    tab.route("**/uji-refresh", lambda r: r.fulfill(content_type="text/html", body=jawaban.pop(0)))
                    tab.goto("https://fasih-sm.bps.go.id/uji-refresh", wait_until="commit")
                    with patch.object(tab, "wait_for_timeout") as tunggu:
                        expect(tab.get_by_role("button", name=tujuan, exact=True)).to_be_visible(timeout=10000)
                        assert tunggu.call_args_list == [call(30000), call(30000)]
                    assert not jawaban
                    tab.unroute("**/uji-refresh")
                    assert tab.get_by_role("button", name=tujuan, exact=True).is_visible()
                    with patch.object(tab, "wait_for_timeout") as tunggu:
                        expect(tab.get_by_role("button", name=tujuan, exact=True)).to_be_visible(timeout=10000)
                        tunggu.assert_not_called()
                page.set_content('<button>Refresh Page</button>')
                with patch.object(page, "wait_for_timeout", side_effect=lambda _: page.set_content('<button>Ringkasan</button>')) as tunggu:
                    expect(page.get_by_role("button", name="Ringkasan", exact=True)).to_be_visible(timeout=10000)
                    tunggu.assert_called_once_with(30000)
                page.set_content('<button hidden onclick="document.body.innerHTML=\'<label>Pekerja laki-laki<input value=1></label>\'">Refresh Page</button>')
                page.evaluate("setTimeout(() => document.querySelector('button').hidden=false, 100)")
                with patch.object(page, "wait_for_timeout") as tunggu:
                    expect(page.get_by_label("Pekerja laki-laki")).to_have_count(1, timeout=10000)
                    tunggu.assert_called_once_with(30000)
                page.set_content('<button onclick="if(this.textContent===\'Refresh Page\')this.textContent=\'CATATAN\';else this.dataset.clicked=\'1\'">Refresh Page</button>')
                with patch.object(page, "wait_for_timeout") as tunggu:
                    page.get_by_role("button", name="CATATAN", exact=True).click(timeout=10000)
                    tunggu.assert_called_once_with(30000)
                assert page.get_by_role("button", name="CATATAN", exact=True).get_attribute("data-clicked") == "1"
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
                with TestCase().assertRaisesRegex(AssertionError, "tidak sesuai"):
                    buka_usaha(page, "USAHA CONTOH <PEMILIK A>")
                page.set_content('<div id="se2026_nested">' + kartu_usaha + '</div><section id="nama" hidden>'
                                 '<div id="nama_usaha_edit#1002">' + wadah_nama + '<div>USAHA CONTOH (PEMILIK B)</div></div></div></section>'
                                 '<div>USAHA CONTOH &lt;PEMILIK A&gt;</div>')
                with TestCase().assertRaisesRegex(AssertionError, "tidak sesuai"):
                    buka_usaha(page, "USAHA CONTOH <PEMILIK A>")
                page.set_content('<div id="se2026_nested">' + kartu_usaha
                                 + '<div data-nested-view="true" onclick="this.dataset.clicked=\'true\'"><span>USAHA CONTOH &lt;PEMILIK A&gt;</span></div></div>')
                buka_usaha(page, "USAHA CONTOH <PEMILIK A>")
                assert page.locator('[data-clicked="true"]').inner_text() == "USAHA CONTOH <PEMILIK A>"
                page.set_content('<div id="se2026_nested"><div data-nested-view="true"><div data-nested-view="true" '
                                 'onclick="this.dataset.clicked=\'true\'"><span>USAHA&nbsp;CONTOH  TIGA</span></div></div></div>')
                buka_usaha(page, "USAHA CONTOH TIGA")
                assert page.locator('[data-clicked="true"]').count() == 1
                page.set_content('<div id="se2026_nested"><div data-nested-view="true" onclick="this.dataset.clicked=\'true\'">USAHA KEMBAR</div>'
                                 '<div data-nested-view="true" onclick="this.dataset.clicked=\'true\'">USAHA KEMBAR</div></div>')
                with TestCase().assertRaisesRegex(AssertionError, "beberapa kartu"):
                    buka_usaha(page, "USAHA KEMBAR")
                assert page.locator('[data-clicked="true"]').count() == 0
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
            page.close()
            page = assignment
            page.bring_to_front()
            assert context.pages == [assignment] and not assignment.is_closed()
            if run == 0:
                root = Path(temp) / "uji-error"
                root.mkdir()
                cfg = dict(excel_file="uji.xlsx", sheet_name="Sheet1", delay_min=0, delay_max=0,
                           batch_min=2, batch_max=2, rest_min=0, rest_max=0, action_min=0, action_max=0,
                           typing_min=1, typing_max=1)
                (root / "config.json").write_text(json.dumps(cfg), encoding="utf-8")
                book = Workbook()
                sheet = book.active
                sheet.title = "Sheet1"
                sheet.append(["id_fasih", "target_usaha", *FIELDS, "remark", "status", "catatan"])
                sheet.append(["macet", "USAHA CONTOH TIGA", *([None] * len(FIELDS)), None, "TODO", "Catatan operator"])
                sheet.append(["lanjut", "USAHA BERIKUTNYA", *([None] * len(FIELDS)), None, "TODO"])
                sheet.append(["ubah", "USAHA BERIKUTNYA", *[0 if k in ("tk_laki_laki", "upah_gaji") else None for k in FIELDS], "Hasil konfirmasi", "TODO"])
                sheet.append(["tetap", "USAHA BERIKUTNYA", *[0 if k in ("tk_laki_laki", "upah_gaji") else None for k in FIELDS], "Tidak disalin", "TODO"])
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
                    '<div id="cek_anomali_button"><input role="switch" aria-checked="true"></div>'
                    + ''.join('<div id="' + key + '#2099"><div class="tw:rounded-lg tw:gap-2"><div>' + label + '</div><input value="'
                              + ('1' if r.request.url.endswith('/ubah') else '0') + '"></div>' + button_note + '</div>'
                              for key, label in (("pekerja", FIELDS["tk_laki_laki"]), ("gaji", FIELDS["upah_gaji"])))
                    + '<section id="editor" hidden><textarea placeholder="Tambah catatan di sini..."></textarea>'
                    '<button onclick="localStorage.setItem(\'remarks_\'+document.body.dataset.field,document.querySelector(\'textarea\').value)">Simpan</button>'
                    '<button onclick="this.parentElement.hidden=true">Dismiss</button></section><button>Kirim</button>')))
                with patch("fasih.__file__", str(root / "fasih.py")), patch("fasih.TIMEOUT", 1000), \
                     patch("fasih.Camoufox", return_value=nullcontext(context)), patch("builtins.input", return_value=""), \
                     patch("fasih.isi_remarks", wraps=isi_remarks) as remarks, patch("fasih.pesan_galat", return_value="Validasi uji"):
                    main()
                    assert remarks.call_count == 2
                    assert all(c.args[2] == "Hasil konfirmasi" for c in remarks.call_args_list)
                book = load_workbook(root / "uji.xlsx")
                data = list(book["Sheet1"].values)
                headers = list(data[0])
                assert data[1][headers.index("status")] == "ERROR"
                note = data[1][headers.index("catatan")]
                assert note.startswith("Catatan operator\nBLOK II: membuka usaha USAHA CONTOH TIGA | AssertionError:")
                assert "Kartu usaha tidak ditemukan" in note
                assert data[2][headers.index("status")] == "NO_CHANGE"
                assert data[3][headers.index("status")] == "ERROR"
                assert data[3][headers.index("keterangan")] == "Validasi uji"
                assert data[4][headers.index("status")] == "NO_CHANGE"
                assert assignment.evaluate("localStorage.getItem('remarks_gaji#2099')") == "Hasil konfirmasi"
                assert assignment.evaluate("localStorage.getItem('remarks_pekerja#2099')") == "Hasil konfirmasi"
                assert context.pages == [assignment]
                book.close()
            context.close()
