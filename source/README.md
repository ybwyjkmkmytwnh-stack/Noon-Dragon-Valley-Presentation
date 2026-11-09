# إعادة البناء
من هذا المجلد ثبّت pptxgenjs@4.0.1 في بيئة Node ثم شغّل `node build.cjs raw.pptx` و`python package_pptx.py raw.pptx final.pptx --fade`. احذف الخيار لإصدار بلا انتقالات. الخط Noto Sans Arabic. يمكن ضبط NOON_PRESENTER وNOON_TEAM محليًا للاعتمادات الخاصة بالحفل. لا تضف القيم إلى المستودع العام.

## إصدار v1.1.0
شغّل `node build.cjs final.pptx` لإخراج بلا تعديل XML. اضبط NOON_SIMPLE_CHART=1 لإخراج رسم الأشكال. NOON_FLOOR_MAP اختياري لنسخة الحفل فقط. استخدم Arial وتحقق من عرضه على جهاز الحفل.
