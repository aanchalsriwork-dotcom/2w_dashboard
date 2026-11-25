"""
ONE-TIME setup for offline translation (Meta NLLB-200 running on CTranslate2).

    pip install ctranslate2 transformers sentencepiece torch
    python setup_nllb.py

What it does
  1. Downloads facebook/nllb-200-distilled-600M from Hugging Face (about 2.5 GB download)
  2. Converts it to a compact int8 CPU model (about 600 MB) in  models/nllb-200-distilled-600M-ct2/
  3. Saves the tokenizer next to it
  4. Runs a real translation test through the same code the news job uses

After it succeeds you can uninstall torch - the news job only needs ctranslate2, transformers (tokenizer
only) and sentencepiece. Run this script again only if you delete the models folder.
"""
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODEL = "facebook/nllb-200-distilled-600M"
OUT = os.path.join(BASE_DIR, "models", "nllb-200-distilled-600M-ct2")


def main():
    try:
        from ctranslate2.converters import TransformersConverter
        from transformers import AutoTokenizer
    except ImportError as error:
        sys.exit(f"Missing package: {error.name}\nRun: pip install ctranslate2 transformers sentencepiece torch")

    if os.path.isdir(OUT) and os.path.exists(os.path.join(OUT, "model.bin")):
        print(f"Model already converted in {OUT} - skipping download/convert.")
    else:
        print(f"Downloading and converting {MODEL} (this takes a few minutes) ...")
        TransformersConverter(MODEL).convert(OUT, quantization="int8", force=True)
        AutoTokenizer.from_pretrained(MODEL).save_pretrained(OUT)
        print(f"Saved model to {OUT}")

    print("Testing the model ...")
    sys.path.insert(0, BASE_DIR)
    from news_fetch_NLLB import NLLBEngine
    engine = NLLBEngine(OUT)
    tests = [("es", "Honda lanza una nueva motocicleta eléctrica en México"),
             ("de", "Neuzulassungen von Motorrädern steigen im ersten Halbjahr"),
             ("th", "ฮอนด้า เปิดตัว มอเตอร์ไซค์ไฟฟ้า รุ่นใหม่"),
             ("ar", "هوندا تطلق دراجة نارية جديدة في مصر"),
             ("vi", "Honda ra mắt xe máy điện mới tại Việt Nam")]
    for lang, text in tests:
        print(f"  [{lang}] {text}\n       -> {engine.translate_batch([text], lang)[0]}")
    print("\nSetup finished. The news job will now translate offline.")


if __name__ == "__main__":
    main()
