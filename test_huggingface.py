#!/usr/bin/env python
"""
Teste la détection des champs (IA + règles).
Usage :  python test_huggingface.py
"""
import os
import sys

import django

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "charge_amendements.settings")
django.setup()

from amendements.huggingface_utils import field_detector  # noqa: E402

ENTETES = [
    "التعديل رقم : 4 الفريق 1 إضافة مادة جديدة مدونة الجمارك والضرائب غير المباشرة المادة 3 البند I الفصل 164 المكرر",
    "التعديل رقم : 2 الفريق 2 الضرائب الداخلية على الاستهلاك المادة 5 البند I الفصل 9",
    "التعديل رقم : 13 الفريق 2 إضافة مادة جديدة المدونة العامة للضرائب المادة 8 البند I المادة 42 المكرر مرّتين",
]


def main():
    print(f"Fournisseur IA : {field_detector.description()}\n")
    ok_ia = False
    for entete in ENTETES:
        valeurs, details = field_detector.detecter(entete)
        print("EN-TÊTE :", entete)
        for champ, d in details.items():
            print(f"   {champ:20s} = {d['value']}   [{d['method']}, {d['confidence']:.2f}]")
            ok_ia = ok_ia or "ia" in d["method"]
        print()
    if field_detector.derniere_erreur:
        print("⚠️  IA indisponible :", field_detector.derniere_erreur)
    print("✅ IA utilisée" if ok_ia else "ℹ️  Seules les règles ont été utilisées (vérifier AI_API_KEY dans .env)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
