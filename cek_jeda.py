from unittest.mock import Mock, call, patch
from decimal import Decimal
from fasih import angka, hitung_mundur, jeda, pola_usaha

assert all(angka(v) is None for v in (None, "", " \t "))
assert all(angka(v) == Decimal(0) for v in (0, 0.0, "0", "0,00"))
assert angka("1.234,50") == Decimal("1234.50")
assert angka(12.5) == Decimal("12.5")

cfg = dict(delay_min=3, delay_max=3, batch_min=2, batch_max=2, rest_min=2, rest_max=2)
sisa, waktu = 2, []
for _ in range(5):
    delay, sisa = jeda(cfg, sisa)
    waktu.append(delay)
assert waktu == [3000, 120000, 3000, 120000, 3000]
assert sisa == 1
assert jeda({**cfg, "batch_min": 1, "batch_max": 1}, 1) == (120000, 1)

page = Mock()
with patch("fasih.monotonic", side_effect=[10, 10, 11, 12, 12.5]), patch("builtins.print") as output:
    hitung_mundur(page, 2500)
    assert page.wait_for_timeout.call_args_list == [call(1000), call(1000), call(500)]
    assert output.call_args_list == [
        call(f"\rIstirahat batch: 00:{s:02d}", end="", flush=True) for s in (3, 2, 1, 0)
    ] + [call(" - lanjut.")]
page.reset_mock()
with patch("fasih.monotonic", return_value=10), patch("builtins.print"):
    hitung_mundur(page, 0)
    page.wait_for_timeout.assert_not_called()

pattern = pola_usaha("USAHA ONLINE (PEMILIK A)")
assert pattern.fullmatch("USAHA ONLINE PRODUK CONTOH (PEMILIK A)")
assert pattern.fullmatch("usaha online (pemilik a)")
assert not pattern.fullmatch("USAHA ONLINE PRODUK CONTOH (PEMILIK B)")
assert not pattern.fullmatch("USAHA ONLINEX (PEMILIK A)")
assert not pattern.fullmatch("USAHA ONLINE (PEMILIK A) CABANG")
assert not pattern.search("TOKO USAHA ONLINE (PEMILIK A)")
assert not pattern.search("USAHA ONLINE (PEMILIK A) CABANG")
assert not pola_usaha("TOKO MAJU").fullmatch("TOKO MAJU JAYA")
assert pola_usaha("USAHA CONTOH <PEMILIK A>").fullmatch("usaha contoh (pemilik a)")
assert pola_usaha("USAHA CONTOH (PEMILIK A)").fullmatch("USAHA CONTOH <PEMILIK A>")
assert pola_usaha("USAHA CONTOH ( PEMILIK CONTOH)").fullmatch("USAHA CONTOH (PEMILIK CONTOH)")
assert not pola_usaha("USAHA CONTOH <PEMILIK A>").fullmatch("USAHA CONTOH <PEMILIK B>")
assert not pola_usaha("USAHA CONTOH <PEMILIK A>").fullmatch("USAHA CONTOH (PEMILIK A>")
pattern = pola_usaha("USAHA CONTOH DUA <PEMILIK CONTOH DUA")
assert pattern.fullmatch("USAHA CONTOH DUA (PEMILIK CONTOH DUA)")
assert pattern.fullmatch("USAHA CONTOH DUA <PEMILIK CONTOH DUA>")
assert pattern.fullmatch("USAHA CONTOH DUA <PEMILIK CONTOH DUA")
assert not pattern.fullmatch("USAHA CONTOH DUA (PEMILIK LAIN)")
assert not pattern.fullmatch("USAHA CONTOH DUA (PEMILIK CONTOH DUA) CABANG")
