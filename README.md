# Lesson PPTX Service

خدمة صغيرة (Node.js + Express + pptxgenjs) تستقبل محتوى الدرس بصيغة JSON
وترجع ملف PowerPoint (.pptx) مصمم فعلياً (تخطيطات متنوعة، ألوان، ظلال،
دعم كامل للعربية RTL) بدل مجرد استبدال نص داخل شريحة واحدة ثابتة.

## نشر الخدمة على Render

1. ارفع هذا المجلد كامل إلى مستودع GitHub جديد (خاص أو عام).
2. من لوحة Render: **New → Web Service** → اربط المستودع.
3. الإعدادات:
   - **Environment:** Node
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Instance Type:** Free يكفي للتجربة
4. بعد النشر، Render رح يعطيك رابط ثابت شكله:
   ```
   https://lesson-pptx-service.onrender.com
   ```
5. جرب أول شي بفتح الرابط بالمتصفح — لازم يرجع:
   ```json
   { "status": "ok", "service": "lesson-pptx-service" }
   ```

## استخدامها من n8n

بدل كل سلسلة (Copy file + Code + Google Slides batchUpdate)، تحتاج بس:

**نود HTTP Request واحد:**
- **Method:** POST
- **URL:** `https://<your-render-url>/generate`
- **Body (JSON):**
```json
{
  "topic": "={{ $json.lesson_data.topic }}",
  "grade": "={{ $json.lesson_data.grade }}",
  "subject": "={{ $json.lesson_data.subject }}",
  "slides": "={{ $json.slides }}"
}
```
- **Response Format:** File / Binary (فعّلها من تبويب Settings بالنود، عشان
  n8n يستقبل ملف الـ pptx كـ binary مو كنص).

الرد بيرجع ملف `.pptx` جاهز مباشرة — تقدر تحطه بنود **Google Drive → Upload**
عشان يترفع تلقائياً لحساب المعلم.

## تعديل التصميم

كل التصميم موجود بملف `pptx-builder.js`:
- `PALETTE` — الألوان الأساسية (غيّرها لألوان مدرستك/تطبيقك).
- `buildTitleSlide`, `buildBulletSlide`, `buildHighlightSlide`,
  `buildSplitSlide`, `buildClosingSlide` — كل دالة تبني نوع شريحة مختلف.
- `pickLayoutBuilder` — المنطق يلي يقرر أي تخطيط يستخدم لكل شريحة (حالياً:
  تخطيط "تمييز" للمحتوى القصير، وتناوب بين تخطيطين للمحتوى الطويل).

## اختبار محلي (اختياري)

```bash
npm install
node test-local.js
```

بيولّد ملف `sample-output.pptx` بمحتوى تجريبي للتأكد إنه كل شي شغال قبل
ما تنشره على Render.
