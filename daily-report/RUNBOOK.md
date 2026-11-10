# PDC Mecca — daily report runbook

Builds the 3-slide daily report from `template.pptx` + one `data_YYYY-MM-DD.json` file:
`python3 build_report.py template.pptx data_YYYY-MM-DD.json output/التقرير_اليومي_DD-MM-YYYY.pptx`
The script prints `MISSING:` with every field that was null/absent (those parts keep the template values).
Use `data_2026-11-09.json` as the reference for the JSON shape.

## Where the data comes from
Gmail (mohamedkamal6492@gmail.com), email from `mohamed.h.kamal@outlook.com`, subject like `تقرير يومي D-M-YYYY`.
Check spam too (`in:anywhere`). The email BODY holds the WhatsApp texts. The attachments
(safety PPTX + المشاعر image) cannot be read through the Gmail tool.

## Field mapping (email text → JSON)
- `date`: Arabic weekday + DD/MM/YYYY of the report date (from the subject / text), e.g. `الثلاثاء 10/11/2026`.
- **SAP block** (`أوامر العمل المسندة في نظام SAP`): مكتمل → `sap.completed`, قيد المراجعة → `sap.in_review`,
  لم يتم البدء → `sap.not_started` (count + amount). الكابلات المؤقتة → `sap_extra.temp_cables`,
  إيصالات المولدات → `sap_extra.generator_receipts`.
- **قسم الإنشاءات block**: التوصيلات/المشاريع/الإجمالي → `construction.assigned_connections/projects/total`;
  إجراء (190/203) → `wo_190_203`; إجراء (155) → `wo_155`; جاهزة (156) → `cert_ready_156`;
  مقبولة (170) → `cert_accepted_170`; (( تم قبول N شهادة)) → `cert_accepted_today`; مرفوضة (171) → `cert_rejected_171`;
  إجراء (230) → `close_230`; بدون غرامة → `close_without_fine`; بغرامة → `close_with_fine`; إجراء (300) → `close_300`;
  الشهادات المؤكدة بغرامات → `fines_certificates`; إجمالي الغرامات → `fines_amount`.
- **متابعة استلام الأصول — أولاً الموقف الإجمالي**: المشاريع/التوصيلات/الحرم/خليص/الليث → `assets.by_entity.projects/connections/haram/khulais/allith`;
  المقبولة → `assets.overall.accepted`; قيد الإجراء → `in_progress`; المتأخرة → `late`; المرفوضة بملاحظات → `rejected_with_notes`.
- **ثانياً موقف طلبات PDC**: إجمالي → `assets.pdc.total`; المقبولة/قيد الإجراء/الجاهزة للإرسال/تحتاج مراجعة/المرفوضة بملاحظات/المتأخرة →
  `accepted/in_progress/ready_to_send/need_review/rejected_with_notes/late`. Values like `--11` or `—127` mean 11 / 127; a bare `—` means null.
- **نظام متابعة الإشراف — معلّقات دورة المرفقات**: per consultant `[عند الاستشاري, عند المقاول]` → `attachments_pending`.
  Consultant names exactly: سمير أبو غلية, الصخرة المشرفة, سعود كونسلت, ابعاد الرؤية, المعماريون.
- **Safety** (`safety`) and **المشاعر** (`mashaer`): come only as attachments. If not provided as text, copy them from the
  most recent data file and report them as NOT UPDATED (with the date they come from).

## Rules
- Never invent numbers. Anything not found → null, and list it in the summary.
- Flag anything that looks inconsistent (e.g. totals that don't add up, values identical to the previous day).
