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
- **Safety** (`safety`) and **المشاعر** (`mashaer`): the user now writes them as free text in the email body
  (any wording/order). Read that text carefully and map it to the same JSON shape as `data_2026-11-09.json`:
  - safety.today: الزيارات (+ PMSS سلامة / CMS / PMSS تنفيذ), الورش, الملاحظات (مغلقة/مفتوحة), المخالفات (سلامة/جودة),
    مبلغ المخالفات (سلامة/جودة), الإيقافات. safety.cumulative: same fields + نسبة الإغلاق.
    `report_label` = `التقرير DD · <Arabic month> YYYY`.
  - observations_list: each safety observation as a separate item, text as written (fix only obvious typos in names).
    quality_observations_list: quality observations if any, else [].
  - violations_list: one object per violation: date, description, class (A/B/…), consultant, contractor, amount.
  - mashaer: total, new_today, replacements_total, cabinets_total, meters_total, stations_total,
    replacements.done/in_progress/awaiting_permits/surveyed, cabinets_done, meters_done, stations_done, source_date.
  Any safety/mashaer field not mentioned → null (do NOT carry forward silently). If a whole section is absent from the
  email, copy it from the most recent data file and report it as NOT UPDATED (with the date it comes from).

## Rules
- Never invent numbers. Anything not found → null, and list it in the summary.
- Flag anything that looks inconsistent (e.g. totals that don't add up, values identical to the previous day).
