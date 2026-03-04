<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>لوحة المتدرب</title>
  <link rel="stylesheet" href="/static/style.css">
</head>
<body>
  <div class="header">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="container">
    <div class="card">
      <h1>لوحة المتدرب</h1>

      <div class="row">
        <div style="flex:1; min-width:260px;">
          <div><b>اسم المتدرب:</b> {{ trainee['اسم المتدرب'] }}</div>
          <div><b>رقم المتدرب:</b> {{ trainee['رقم المتدرب'] }}</div>
          <div><b>التخصص:</b> {{ trainee['التخصص'] }}</div>
          <div><b>البرنامج:</b> {{ trainee.get('برنامج','') }}</div>
        </div>

        <div style="flex:1; min-width:260px;">
          <div class="badge">إجمالي فرص تخصصك: {{ total_spec }}</div>
          <div style="margin-top:8px;" class="badge">المتبقي الآن: {{ remaining_spec }}</div>
        </div>
      </div>

      <hr style="margin:18px 0; border:0; border-top:1px solid #eee;">

      {% if assigned %}
        <div class="success">
          تم الحجز مسبقاً ✅
        </div>
        <div style="margin-top:8px;">
          <b>الجهة المختارة:</b> {{ assigned['entity'] }}
        </div>

        <div style="margin-top:14px;">
          <a href="/download/{{ trainee['رقم المتدرب'] }}">
            <button type="button" class="secondary">تحميل PDF</button>
          </a>
        </div>

      {% else %}
        <form method="post" action="/assign/{{ trainee['رقم المتدرب'] }}">
          <label><b>اختر جهة التدريب</b> <span class="small">(يظهر المتبقي لكل جهة)</span></label>
          <div class="row" style="margin-top:8px;">
            <div class="input">
              <select name="entity" required>
                <option value="">— اختر —</option>
                {% for item in entities %}
                  <option value="{{ item['name'] }}">{{ item['name'] }} (المتبقي: {{ item['remaining'] }})</option>
                {% endfor %}
              </select>
            </div>
            <div style="align-self:end;">
              <button type="submit">حجز + طباعة PDF</button>
            </div>
          </div>

          {% if error %}
            <div class="error">{{ error }}</div>
          {% endif %}
          {% if msg %}
            <div class="success">{{ msg }}</div>
          {% endif %}
        </form>
      {% endif %}
    </div>
  </div>
</body>
</html>
