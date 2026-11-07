#!/usr/bin/env python3
import json
import os
from pathlib import Path

def get_system_documents_dir():
    if "APPIMAGE" in os.environ:
        return Path(os.environ["APPIMAGE"]).parent
    else:
        return Path(os.path.realpath(__file__)).parent

PROJE_ANA_DIZINI = get_system_documents_dir()
LANGS_DIR = PROJE_ANA_DIZINI / "langs"

# Hafızada aktif dili ve kelime önbelleğini (cache) tutuyoruz
AKTIF_DIL = "en"
ILK_KURULUM_METINLERI = {}

def dili_yukle(lang_kodu):
    """Belirtilen dil koduna ait JSON dosyasını dinamik olarak okur."""
    global ILK_KURULUM_METINLERI
    json_path = LANGS_DIR / f"{lang_kodu}.json"

    if json_path.exists():
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                ILK_KURULUM_METINLERI = json.load(f)
                return True
        except Exception as e:
            print(f"[Dil Yukleme Hatasi] JSON okunurken hata olustu: {e}")

    # Eğer JSON bulunamaz veya bozuksa boşalt ki t() fonksiyonu anahtarları döndürsün
    ILK_KURULUM_METINLERI = {}
    return False

def t(anahtar):
    """Anahtar kelimeyi aktif dilde döndürür, bulamazsa anahtarın kendisini verir."""
    return ILK_KURULUM_METINLERI.get(anahtar, anahtar)

def set_lang(lang_kodu):
    """Aktif dili değiştirir ve ilgili sözlüğü yükler."""
    global AKTIF_DIL
    AKTIF_DIL = lang_kodu
    dili_yukle(lang_kodu)

def get_all_translations_for_key(anahtar):
    """
    Ana penceredeki arama agacini temizlerken tum dillerdeki karsiligini kontrol etmek icin yardimci fonksiyon.
    Eski kodda METINLER.values() taraması yapan mantığın dinamik halidir.
    """
    karsiliklar = []
    if LANGS_DIR.exists():
        for dosya in LANGS_DIR.glob("*.json"):
            try:
                with open(dosya, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if anahtar in data:
                        karsiliklar.append(data[anahtar])
            except Exception:
                pass
    return karsiliklar if karsiliklar else [anahtar]

# Uygulama ilk açıldığında varsayılan dili hafızaya yükle
set_lang(AKTIF_DIL)
