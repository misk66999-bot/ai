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

## تفعيل الصور الحقيقية (Pexels)

1. سجّل حساب مجاني على https://www.pexels.com/api/ واخذ الـ API Key.
2. بلوحة Render، روح لمشروعك → **Environment** → أضف متغيّر بيئة جديد:
   - **Key:** `PEXELS_API_KEY`
   - **Value:** المفتاح يلي أخذته
3. احفظ — Render بيعيد نشر الخدمة تلقائياً.

**ملاحظة مهمة:** بحث Pexels أدق بكثير بالإنجليزي. لو بدك نتائج صور أفضل،
عدّل برومبت Groq ليرجع حقل إضافي `visual_suggestion_en` (وصف الصورة
بالإنجليزي، كلمتين-ثلاثة، مثل `"electric circuit diagram"`) جنب
`visual_suggestion` العربي. الخدمة تستخدم `visual_suggestion_en` تلقائياً
لو موجود، وإلا تستخدم النص العربي كما هو.

لو ما ضفت مفتاح Pexels أصلاً، الخدمة تشتغل عادي بدون صور (بس التصميم
النصي الملوّن السابق) — ما في أي كسر.

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
