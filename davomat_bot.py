# -*- coding: utf-8 -*-
"""
DAVOMAT BOTI v3.0 - institut malaka oshirish kursi uchun
 - Kontakt orqali ro'yxatdan o'tish + doimiy "Kirish" tugmasi
 - Darslar jadvali (para, fan, o'qituvchi, xona, bino, kun, vaqt)
 - Har para boshlanishida avtomatik xabar yuborish
 - Davomat olish (istalgan sana), hisobot, oylik foiz, Excel
 - Ota-onaga xabar, o'qituvchi va foydalanuvchini guruhga biriktirish

O'rnatish:  pip install aiogram openpyxl
Ishga tushirish: python davomat_bot.py
"""
import sqlite3, logging, datetime, asyncio, os
try:
    from aiohttp import web
    HAS_WEB = True
except ImportError:
    HAS_WEB = False
    web = None
from aiogram import Bot, Dispatcher, F
from aiogram.types import (Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton,
    FSInputFile, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove, WebAppInfo)
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from openpyxl import Workbook

TOKEN = os.getenv("TOKEN", "BU_YERGA_BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

# Render avtomatik beradi; mahalliy sinov uchun PUBLIC_URL="https://..." yozing
PUBLIC_URL = os.getenv("RENDER_EXTERNAL_URL", os.getenv("PUBLIC_URL", ""))

# ===== MINI APP =====
# HTTPS manzil (cloudflared/ngrok yoki server domeni)
WEBAPP_URL = "https://SIZNING-HTTPS-MANZIL"
PORT = int(os.getenv("PORT", 8080))
RUN_WEB = True    # Web-server har doim ishlaydi (Render webhook + Mini App)
BASE = os.path.dirname(os.path.abspath(__file__))

INDEX_HTML = """<!DOCTYPE html>
<html lang="uz">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>Navoiy Davomat</title>
<script src="https://telegram.org/js/telegram-web-app.js"></script>
<style>
  * { margin:0; padding:0; box-sizing:border-box; -webkit-tap-highlight-color:transparent; }
  body { background:#0e0e10; color:#e8e8ea; font-family:-apple-system,'Segoe UI',Roboto,sans-serif; padding-bottom:80px; }
  .hidden { display:none !important; }

  /* ====== header profile ====== */
  .profile { display:flex; gap:14px; padding:16px; align-items:center; }
  .avatar { width:64px; height:64px; border-radius:50%; background:#2a2a2e; display:flex; align-items:center; justify-content:center; font-size:26px; font-weight:700; color:#3a86ff; flex-shrink:0; overflow:hidden; }
  .avatar img { width:100%; height:100%; object-fit:cover; }
  .p-name { font-size:19px; font-weight:700; }
  .p-sub { font-size:13px; color:#9a9aa0; margin-top:3px; line-height:1.35; }
  .pct-badge { margin-left:auto; background:rgba(46,204,113,.15); color:#2ecc71; font-weight:700; padding:8px 14px; border-radius:20px; font-size:15px; white-space:nowrap; }

  .card { background:#1c1c1f; border-radius:18px; margin:10px 14px; padding:16px; }
  .card-title { font-size:13px; color:#9a9aa0; text-transform:uppercase; letter-spacing:.5px; margin-bottom:10px; font-weight:600; }

  /* time + para */
  .time-row { display:flex; justify-content:space-between; align-items:center; }
  .clock { font-size:42px; font-weight:800; }
  .date { color:#9a9aa0; font-size:14px; margin-top:2px; }
  .para-box { text-align:right; }
  .para-label { font-size:11px; color:#9a9aa0; text-transform:uppercase; }
  .para-val { font-size:17px; font-weight:700; color:#3a86ff; }

  /* location */
  .loc { display:flex; align-items:center; gap:8px; font-size:14px; padding:12px 14px; background:#1c1c1f; border-radius:14px; margin:10px 14px; }
  .dot { width:9px; height:9px; border-radius:50%; background:#e74c3c; box-shadow:0 0 8px #e74c3c; }
  .loc-txt { flex:1; }
  .loc-err { color:#ff6b6b; font-weight:600; }
  .loc-sub { color:#7a7a80; font-size:12px; }

  /* stats */
  .stats { display:flex; gap:10px; margin:10px 14px; }
  .stat { flex:1; background:#1c1c1f; border-radius:16px; padding:14px 12px; text-align:center; }
  .stat .n { font-size:24px; font-weight:800; }
  .stat .l { font-size:12px; color:#9a9aa0; margin-top:2px; }
  .g { color:#2ecc71; } .r { color:#ff6b6b; } .b { color:#3a86ff; }

  /* buttons */
  .btn-row { display:flex; gap:10px; margin:10px 14px; }
  .btn { flex:1; border:none; border-radius:16px; padding:16px 10px; font-size:16px; font-weight:700; color:#fff; display:flex; flex-direction:column; align-items:center; gap:6px; cursor:pointer; }
  .btn:active { transform:scale(.97); }
  .btn-in { background:linear-gradient(135deg,#f39c12,#e67e22); }
  .btn-out { background:linear-gradient(135deg,#5a5a60,#3a3a3e); color:#bbb; }
  .btn .ico { font-size:26px; }

  /* schedule */
  .lesson { background:#232327; border-radius:14px; padding:14px; margin-top:10px; }
  .lesson-top { display:flex; justify-content:space-between; font-size:12px; color:#9a9aa0; margin-bottom:6px; }
  .lesson-name { font-size:16px; font-weight:700; margin-bottom:6px; }
  .lesson-meta { font-size:13px; color:#b5b5bb; display:flex; flex-direction:column; gap:3px; }
  .lesson-status { margin-top:10px; text-align:center; background:#2a2a2e; border-radius:10px; padding:8px; font-size:14px; color:#9a9aa0; }

  /* profile rows */
  .row { display:flex; justify-content:space-between; align-items:center; padding:13px 0; border-bottom:1px solid #2a2a2e; font-size:15px; }
  .row:last-child { border-bottom:none; }
  .row .k { color:#9a9aa0; } .row .v { font-weight:600; text-align:right; max-width:60%; }

  /* bottom nav */
  .nav { position:fixed; bottom:0; left:0; right:0; background:#161618; border-top:1px solid #26262a; display:flex; padding:6px 2px calc(6px + env(safe-area-inset-bottom)); z-index:100; }
  .nav-item { flex:1; text-align:center; padding:5px 0; cursor:pointer; color:#7a7a80; }
  .nav-item.active { color:#3a86ff; }
  .nav-item .ni { font-size:20px; }
  .nav-item .nl { font-size:10px; margin-top:2px; }

  .soon { text-align:center; color:#5a5a60; padding:60px 20px; font-size:15px; }
  .skeleton { background:#2a2a2e; border-radius:8px; color:transparent; }
</style>
</head>
<body>

<!-- ============ ASOSIY ============ -->
<div id="tab-asosiy">
  <div class="profile">
    <div class="avatar" id="avatar">?</div>
    <div>
      <div class="p-name" id="pname">Yuklanmoqda...</div>
      <div class="p-sub" id="psub"></div>
    </div>
    <div class="pct-badge" id="pct">—%</div>
  </div>

  <div class="card">
    <div class="time-row">
      <div><div class="clock" id="clock">--:--</div><div class="date" id="dateline"></div></div>
      <div class="para-box"><div class="para-label">Hozirgi dars</div><div class="para-val" id="paraval">—</div></div>
    </div>
  </div>

  <div class="loc">
    <div class="dot"></div>
    <div class="loc-txt"><div class="loc-err">\u23f3 Lokatsiya aniqlanmoqda...</div><div class="loc-sub" id="locsub"></div></div>
  </div>

  <div class="stats">
    <div class="stat"><div class="n g" id="st-keldi">0</div><div class="l">Qatnashgan</div></div>
    <div class="stat"><div class="n r" id="st-nb">0</div><div class="l">NB soat</div></div>
    <div class="stat"><div class="n b" id="st-jami">0</div><div class="l">Jami soat</div></div>
  </div>

  <div class="btn-row">
    <button class="btn btn-in" id="btn-in"><span class="ico">📍</span>Kirish</button>
    <button class="btn btn-out" id="btn-out"><span class="ico">↩️</span>Chiqish</button>
  </div>

  <div class="card">
    <div class="card-title">Dars jadvali</div>
    <div id="lessons"><div class="soon">Darslar yuklanmoqda...</div></div>
  </div>
</div>

<!-- ============ MONITORING ============ -->
<div id="tab-monitoring" class="hidden"><div class="soon">📊 Monitoring — 3-bosqichda qo'shiladi<br><br>Grafiklar, foizlar, kunlik tarix...</div></div>

<!-- ============ AI ============ -->
<div id="tab-ai" class="hidden"><div class="soon">✨ AI yordamchi — 5-bosqichda qo'shiladi</div></div>

<!-- ============ FANLAR ============ -->
<div id="tab-fanlar" class="hidden"><div class="soon">📚 Fanlar jurnali — 4-bosqichda qo'shiladi</div></div>

<!-- ============ MALAKAM ============ -->
<div id="tab-malakam" class="hidden"><div class="soon">🌐 Malakam (OneID) — 6-bosqichda qo'shiladi</div></div>

<!-- ============ PROFIL ============ -->
<div id="tab-profil" class="hidden">
  <div class="card" style="text-align:center; padding:24px 16px;">
    <div class="avatar" style="margin:0 auto 12px;" id="avatar2">?</div>
    <div class="p-name" id="pname2" style="font-size:22px;">—</div>
    <div class="p-sub" id="psub2"></div>
  </div>
  <div class="card">
    <div class="card-title">Shaxsiy ma'lumotlar</div>
    <div class="row"><span class="k">Telefon</span><span class="v" id="pr-phone">—</span></div>
    <div class="row"><span class="k">Tashkilot</span><span class="v">Navoiy viloyati pedagogik mahorat markazi</span></div>
    <div class="row"><span class="k">Guruh</span><span class="v" id="pr-group">—</span></div>
    <div class="row"><span class="k">Ishtirok</span><span class="v" id="pr-pct">—</span></div>
  </div>
</div>

<!-- ============ NAV ============ -->
<div class="nav">
  <div class="nav-item active" data-tab="asosiy"><div class="ni">🏠</div><div class="nl">Asosiy</div></div>
  <div class="nav-item" data-tab="monitoring"><div class="ni">📊</div><div class="nl">Monitoring</div></div>
  <div class="nav-item" data-tab="ai"><div class="ni">✨</div><div class="nl">AI</div></div>
  <div class="nav-item" data-tab="fanlar"><div class="ni">📖</div><div class="nl">Fanlar jurnali</div></div>
  <div class="nav-item" data-tab="malakam"><div class="ni">🌐</div><div class="nl">Malakam</div></div>
  <div class="nav-item" data-tab="profil"><div class="ni">👤</div><div class="nl">Profil</div></div>
</div>

<script>
const tg = window.Telegram.WebApp;
tg.ready(); tg.expand();
tg.setHeaderColor('#0e0e10'); tg.setBackgroundColor('#0e0e10');

const USER = tg.initDataUnsafe?.user || null;
const API = '';

/* ---- tab switching ---- */
document.querySelectorAll('.nav-item').forEach(el => {
  el.onclick = () => {
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    document.querySelectorAll('[id^=tab-]').forEach(t => t.classList.add('hidden'));
    el.classList.add('active');
    document.getElementById('tab-' + el.dataset.tab).classList.remove('hidden');
    window.scrollTo(0, 0);
  };
});

/* ---- clock ---- */
const WD = ['Yakshanba','Dushanba','Seshanba','Chorshanba','Payshanba','Juma','Shanba'];
const WD_S = ['Ya','Du','Se','Cho','Pa','Ju','Sha'];
function tick() {
  const d = new Date();
  document.getElementById('clock').textContent =
    String(d.getHours()).padStart(2,'0') + ':' + String(d.getMinutes()).padStart(2,'0');
  document.getElementById('dateline').textContent = WD[d.getDay()] + ', ' +
    d.getDate() + ' ' + ['yan','fev','mar','apr','may','iyun','iyul','avg','sen','okt','noy','dek'][d.getMonth()];
}
tick(); setInterval(tick, 10000);

/* ---- load user data ---- */
async function loadMe() {
  if (!USER) { document.getElementById('pname').textContent = 'Telegram ochilmadi'; return; }
  try {
    const r = await fetch(API + '/api/me?id=' + USER.id);
    const d = await r.json();
    if (d.error) {
      document.getElementById('pname').textContent = '⚠️ Ro\'yxatdan o\'tmagansiz';
      document.getElementById('psub').textContent = '';
      document.getElementById('lessons').innerHTML =
        '<div style="text-align:center;padding:30px 16px;color:#9a9aa0;font-size:15px;line-height:1.7">' +
        'Botga qayting va<br><b style="color:#3a86ff">«Ro\'yxatdan o\'ting»</b> tugmasini bosing,' +
        '<br>keyin «Kirish» ni qayta oching</div>';
      document.getElementById('pct').textContent = '—';
      return;
    }
    document.getElementById('pname').textContent = d.name;
    document.getElementById('psub').textContent = (d.spec || '') + (d.group ? ' · ' + d.group : '');
    document.getElementById('pct').textContent = d.pct + '%';
    document.getElementById('st-keldi').textContent = d.keldi;
    document.getElementById('st-nb').textContent = d.nb;
    document.getElementById('st-jami').textContent = d.jami;
    document.getElementById('pr-phone').textContent = d.phone || '—';
    document.getElementById('pr-group').textContent = d.group || '—';
    document.getElementById('pr-pct').textContent = d.pct + '% (' + d.keldi + '/' + d.jami + ')';
    document.getElementById('pname2').textContent = d.name;
    document.getElementById('psub2').textContent = d.spec || '';
    const ini = d.name.split(' ').map(w=>w[0]).slice(0,2).join('');
    document.getElementById('avatar').textContent = ini;
    document.getElementById('avatar2').textContent = ini;
    // lessons
    if (d.lessons && d.lessons.length) {
      document.getElementById('lessons').innerHTML = d.lessons.map(l => `
        <div class="lesson">
          <div class="lesson-top"><span>${l.para}-PARA</span><span>${l.time}</span></div>
          <div class="lesson-name">${l.fan}</div>
          <div class="lesson-meta">
            <span>👤 ${l.oqituvchi}</span>
            <span>📍 ${l.bino}, ${l.xona}</span>
          </div>
          <div class="lesson-status">🕐 ${l.status}</div>
        </div>`).join('');
      // current para countdown
      const now = new Date();
      const cur = d.lessons.find(l => {
        const [h,m] = l.time.split(':').map(Number);
        const st = new Date(); st.setHours(h,m,0,0);
        return st > now;
      });
      if (cur) {
        const [h,m] = cur.time.split(':').map(Number);
        const st = new Date(); st.setHours(h,m,0,0);
        const diff = Math.floor((st - now)/60000);
        document.getElementById('paraval').textContent = diff + "' → " + cur.para + "-para";
      } else document.getElementById('paraval').textContent = 'Dars yo‘q';
    } else {
      document.getElementById('lessons').innerHTML = '<div class="soon">Bugun dars yo‘q</div>';
      document.getElementById('paraval').textContent = '—';
    }
  } catch(e) { document.getElementById('pname').textContent = 'Server bilan aloqa yo‘q'; }
}
loadMe();

/* ---- location: server bilan tekshirish ---- */
let MY_POS = null;
function fmtDist(m) { return m >= 1000 ? (m/1000).toFixed(1) + ' km' : m + ' m'; }
function checkLocation() {
  if (!navigator.geolocation) { locErr('Brauzer lokatsiyani qo\'llamaydi'); return; }
  locWait();
  navigator.geolocation.getCurrentPosition(async p => {
    MY_POS = p.coords;
    if (!USER) { locErr('Telegram foydalanuvchisi topilmadi'); return; }
    try {
      const r = await fetch(API + '/api/location?id=' + USER.id +
        '&lat=' + p.coords.latitude + '&lon=' + p.coords.longitude);
      const d = await r.json();
      INSIDE_NAME = d.inside;
      if (d.inside) {
        locOk('\u2705 ' + d.inside + ' ichidasiz');
      } else {
        locErr('\u2744 Hech bir bino ichida emassiz');
      }
      document.getElementById('locsub').textContent =
        d.buildings.map(b => b.name + ': ' + fmtDist(b.dist)).join(' · ');
    } catch(e) { locErr('Server bilan aloqa yo\'q'); }
  }, () => locErr('\u26a0\ufe0f Lokatsiya ruxsat berilmadi. Sozlamalardan yoqing!'),
     { enableHighAccuracy:true, timeout:15000, maximumAge:30000 });
}
function locWait() {
  const el = document.querySelector('.loc-err');
  el.textContent = '\u23f3 Lokatsiya aniqlanmoqda...'; el.style.color = '#9a9aa0';
}
function locOk(txt) {
  const el = document.querySelector('.loc-err');
  el.textContent = txt; el.style.color = '#2ecc71';
}
function locErr(txt) {
  const el = document.querySelector('.loc-err');
  el.textContent = txt; el.style.color = '#ff6b6b';
}
checkLocation();

/* ---- Kirish / Chiqish: lokatsiya + QR skaner ---- */
let INSIDE_NAME = null;
document.getElementById('btn-in').onclick = () => startScan('in');
document.getElementById('btn-out').onclick = () => startScan('out');

async function startScan(mode) {
  if (!USER) { tg.showAlert('Telegram ichida oching'); return; }
  if (!MY_POS) { tg.showAlert('\u26a0\ufe0f Lokatsiya aniqlanmagan. Sahifani yangilang.'); return; }
  // 1) lokatsiya tekshiruvi
  try {
    const r = await fetch(API + '/api/location?id=' + USER.id +
      '&lat=' + MY_POS.latitude + '&lon=' + MY_POS.longitude);
    const d = await r.json();
    INSIDE_NAME = d.inside;
    if (!d.inside) {
      tg.showAlert('\u274c Siz hech bir bino ichida emassiz!\nDavomat uchun markazga keling.');
      return;
    }
  } catch(e) { tg.showAlert('Server bilan aloqa yo\'q'); return; }
  // 2) QR skaner ochish
  const word = mode === 'in' ? 'KIRISH' : 'CHIQISH';
  tg.showScanQrPopup({text: word + ' uchun QR kodni skanerlang'}, async (qr) => {
    tg.closeScanQrPopup();
    tg.MainButton.show();
    tg.MainButton.setText('\u23f3 Tekshirilmoqda...');
    try {
      const r = await fetch(API + '/api/scan', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({id: USER.id, qr: qr, mode: mode,
                              lat: MY_POS.latitude, lon: MY_POS.longitude})
      });
      const d = await r.json();
      tg.MainButton.hide();
      if (d.ok) { tg.showAlert('\u2705 ' + d.msg); checkLocation(); }
      else { tg.showAlert('\u274c ' + (d.error || 'Xato')); }
    } catch(e) { tg.MainButton.hide(); tg.showAlert('Server bilan aloqa yo\'q'); }
    return true;
  });
}
</script>
</body>
</html>"""


# Salomlashuv matni:
# "Assalomu alaykum, {ism}!" + INSTITUT + ga xush kelibsiz.
INSTITUT = "Navoiy viloyati pedagogik mahorat markaziga"
KURS = ""   # kerak bo'lsa qo'shimcha matn, bo'sh qoldirish mumkin

logging.basicConfig(level=logging.INFO)
bot = Bot(token=TOKEN)
dp = Dispatcher(storage=MemoryStorage())
db = sqlite3.connect("davomat.db", check_same_thread=False)
db.execute("""CREATE TABLE IF NOT EXISTS users(
    tg_id INTEGER PRIMARY KEY, name TEXT, role TEXT, group_id INTEGER, phone TEXT)""")
db.execute("""CREATE TABLE IF NOT EXISTS groups_(
    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT)""")
db.execute("""CREATE TABLE IF NOT EXISTS students(
    id INTEGER PRIMARY KEY AUTOINCREMENT, group_id INTEGER,
    full_name TEXT, parent_id INTEGER)""")
db.execute("""CREATE TABLE IF NOT EXISTS attendance(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER, date TEXT, status TEXT,
    UNIQUE(student_id, date))""")
db.execute("""CREATE TABLE IF NOT EXISTS buildings(
    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT,
    lat REAL, lon REAL, radius_m INTEGER)""")
db.execute("""CREATE TABLE IF NOT EXISTS checkins(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER, lesson_id INTEGER, action TEXT,
    cdate TEXT, ctime TEXT,
    UNIQUE(user_id, lesson_id, action, cdate))""")
db.execute("""CREATE TABLE IF NOT EXISTS lessons(
    id INTEGER PRIMARY KEY AUTOINCREMENT, group_id INTEGER,
    para_num INTEGER, fan TEXT, oqituvchi TEXT, xona TEXT, bino TEXT,
    weekday INTEGER, start_time TEXT)""")

class AddGroup(StatesGroup):    name = State()
class AddStudent(StatesGroup):  name = State(); group = State(); parent = State()
class TakeDate(StatesGroup):    group_id = State(); date = State()
class PercentSt(StatesGroup):   group_id = State(); month = State()
class AssignT(StatesGroup):     teacher = State(); group = State()
class AssignU(StatesGroup):     user = State(); group = State()
class MakeQR(StatesGroup):
    group = State(); lesson = State()

class AddBuilding(StatesGroup):
    name = State(); lat = State(); lon = State(); radius = State()

class AddLesson(StatesGroup):
    group = State(); para = State(); fan = State(); oqituvchi = State()
    xona = State(); bino = State(); weekday = State(); time = State()

def kb(items, row=2):
    rows = [items[i:i+row] for i in range(0, len(items), row)]
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=t, callback_data=c) for t, c in r] for r in rows])

# Mini App URL endi BotFatherda saqlanadi (/mybots -> Menu Button)
# Shuning uchun "Kirish" tugmasi faqat menyuni ochadi
KIRISH_KB = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="🔑 Kirish")]], resize_keyboard=True)

# ====== MINI APP WEB SERVER (faqat lokal rejimda) ======
# ===== XATO TUZATISH YORDAMCHISI (ro'yxatdan o'tgan) =====
import traceback
async def error_middleware(request, handler):
    try:
        return await handler(request)
    except Exception:
        tb = traceback.format_exc()
        logging.error(f"API XATO: {tb}")
        return web.json_response({"xato": tb[-800:]}, status=500)

app = web.Application(middlewares=[error_middleware])

async def index(request):
    path = os.path.join(BASE, "index.html")
    if os.path.exists(path):
        return web.FileResponse(path)
    return web.Response(text=INDEX_HTML, content_type="text/html")

async def api_me(request):
  try:
    tid = int(request.query.get("id", 0) or 0)
    u = db.execute("SELECT name,phone,group_id FROM users WHERE tg_id=?", (tid,)).fetchone()
    if not u: return web.json_response({"error": "not found"}, status=404)
    name, phone, gid = u
    gname = db.execute("SELECT name FROM groups_ WHERE id=?", (gid,)).fetchone() if gid else None
    gname = gname[0] if gname else None
    keldi = nb = jami = 0
    if gid:
        st = dict(db.execute("""SELECT a.status, COUNT(*) FROM attendance a
            JOIN students s ON s.id=a.student_id
            WHERE s.group_id=? GROUP BY a.status""", (gid,)).fetchall())
        keldi = st.get("keldi", 0); nb = st.get("kelmadi", 0) + st.get("sababli", 0)
        jami = sum(st.values())
    pct = round(keldi / jami * 100) if jami else 0
    wd = datetime.datetime.now().weekday()
    lessons = [{"para": p, "fan": f, "oqituvchi": o, "xona": x, "bino": b, "time": t,
                "status": "Boshlanmagan"}
               for p, f, o, x, b, t in db.execute(
                   """SELECT para_num,fan,oqituvchi,xona,bino,start_time
                      FROM lessons WHERE group_id=? AND weekday=? ORDER BY start_time""",
                   (gid, wd)).fetchall()] if gid else []
    return web.json_response({"name": name, "phone": phone, "group": gname,
        "spec": "Kursant", "keldi": keldi, "nb": nb, "jami": jami,
        "pct": pct, "lessons": lessons})
  except Exception:
    import traceback as _tb
    return web.json_response({"XATO": _tb.format_exc()[-800:]}, status=200)

if HAS_WEB:
    app.router.add_get("/", index)

async def api_location(request):
  try:
    lat = float(request.query.get("lat", 0)); lon = float(request.query.get("lon", 0))
    inside = None; res = []
    for bid, name, blat, blon, rad in db.execute(
            "SELECT id,name,lat,lon,radius_m FROM buildings").fetchall():
        d = haversine_m(lat, lon, blat, blon)
        res.append({"name": name, "dist": round(d), "radius": rad, "inside": d <= rad})
        if d <= rad and inside is None:
            inside = name
    return web.json_response({"inside": inside, "buildings": res})
  except Exception:
    import traceback as _tb
    return web.json_response({"XATO": _tb.format_exc()[-800:]}, status=200)

async def api_scan(request):
    import json as _json
    try:
        data = await request.json()
        tid = int(data.get("id", 0)); qr = str(data.get("qr", ""))
        mode = str(data.get("mode", "in"))
        lat = float(data.get("lat", 0)); lon = float(data.get("lon", 0))
    except Exception:
        return web.json_response({"error": "So'rov noto'g'ri"}, status=400)
    u = db.execute("SELECT group_id FROM users WHERE tg_id=?", (tid,)).fetchone()
    if not u: return web.json_response({"error": "Avval ro'yxatdan o'ting!"}, status=403)
    try:
        q = _json.loads(qr)
        lid = int(q["l"]); qdate = str(q["d"])
    except Exception:
        return web.json_response({"error": "QR kod noto'g'ri! Ruxsat etilgan QR emas."}, status=400)
    les = db.execute("SELECT group_id,para_num,bino,fan,start_time FROM lessons WHERE id=?", (lid,)).fetchone()
    if not les: return web.json_response({"error": "Dars topilmadi"}, status=404)
    if u[0] != les[0]:
        return web.json_response({"error": "Bu QR sizning guruhingizga tegishli emas!"}, status=403)
    today = datetime.date.today().isoformat()
    if qdate != today:
        return web.json_response({"error": f"QR boshqa kunga tegishli ({qdate}). Bugun: {today}"}, status=403)
    b = db.execute("SELECT lat,lon,radius_m FROM buildings WHERE name=?", (les[2],)).fetchone()
    if b:
        dist = haversine_m(lat, lon, b[0], b[1])
        if dist > b[2]:
            return web.json_response(
                {"error": f"Siz {les[2]} ichida emassiz! ({round(dist)} m uzoqda)"}, status=403)
    dup = db.execute("SELECT 1 FROM checkins WHERE user_id=? AND lesson_id=? AND action=? AND cdate=?",
                     (tid, lid, mode, today)).fetchone()
    if dup:
        word = "kirish" if mode == "in" else "chiqish"
        return web.json_response({"error": f"{word.capitalize()} allaqachon qayd etilgan!"}, status=409)
    now = datetime.datetime.now()
    db.execute("INSERT INTO checkins(user_id,lesson_id,action,cdate,ctime) VALUES(?,?,?,?,?)",
               (tid, lid, mode, today, now.strftime("%H:%M")))
    db.commit()
    word = "Kirish" if mode == "in" else "Chiqish"
    return web.json_response({"ok": True,
        "msg": f"{word} qayd etildi! \u2705 {les[3]} ({les[1]}-para), vaqt: {now.strftime('%H:%M')}"})

async def api_debug(request):
    steps = {"version": "diag-2"}
    for tbl in ["users","groups_","students","attendance","buildings","checkins","lessons"]:
        try:
            steps[tbl] = db.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
        except Exception as e:
            steps[tbl + "_XATO"] = str(e)
    try:
        u = db.execute("SELECT name,phone,group_id FROM users WHERE tg_id=7079998283").fetchone()
        steps["admin_bor"] = bool(u)
        if u: steps["admin_gid"] = u[2]
    except Exception as e:
        steps["admin_XATO"] = str(e)
    return web.json_response(steps)

app.router.add_get("/api/debug", api_debug)
app.router.add_post("/api/scan", api_scan)
app.router.add_get("/api/location", api_location)
app.router.add_get("/api/me", api_me)

async def api_debug(request):
    steps = {"version": "diag-1"}
    for tbl in ["users", "groups_", "students", "attendance",
                "buildings", "checkins", "lessons"]:
        try:
            steps[tbl] = db.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
        except Exception as e:
            steps[tbl + "_XATO"] = str(e)
    try:
        u = db.execute("SELECT name,phone,group_id FROM users WHERE tg_id=7079998283").fetchone()
        steps["admin_bor"] = bool(u)
        if u: steps["admin_gid"] = u[2]
    except Exception as e:
        steps["admin_XATO"] = str(e)
    return web.json_response(steps)

app.router.add_get("/api/debug", api_debug)

WEEKDAYS = ["Du", "Se", "Cho", "Pa", "Ju", "Sha", "Ya"]

def role_of(tg_id):
    u = db.execute("SELECT role FROM users WHERE tg_id=?", (tg_id,)).fetchone()
    return u[0] if u else "teacher"

def group_of(tg_id):
    u = db.execute("SELECT group_id FROM users WHERE tg_id=?", (tg_id,)).fetchone()
    return u[0] if u else None

def greeting(name):
    extra = f"\n{KURS}" if KURS else ""
    return (f"Assalomu alaykum, {name}!\n\n{INSTITUT}{extra} xush kelibsiz.\n\n"
            f"Tizimga kirish uchun ro'yxatdan o'ting:")

# ================= ro'yxatdan o'tish =================
@dp.message(CommandStart())
async def start(m: Message):
    u = db.execute("SELECT role FROM users WHERE tg_id=?", (m.from_user.id,)).fetchone()
    if not u:
        await m.answer(greeting(m.from_user.first_name), reply_markup=CONTACT_KB())
        return
    await m.answer(f"Xush kelibsiz, {m.from_user.first_name}!", reply_markup=KIRISH_KB)

def CONTACT_KB():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="Ro'yxatdan o'ting", request_contact=True)]],
        resize_keyboard=True)

@dp.message(F.contact)
async def register(m: Message):
    if db.execute("SELECT 1 FROM users WHERE tg_id=?", (m.from_user.id,)).fetchone():
        await m.answer("Siz allaqachon ro'yxatdan o'tgansiz.", reply_markup=KIRISH_KB)
        await menu(m.answer, role_of(m.from_user.id)); return
    role = "admin" if m.from_user.id == ADMIN_ID else "teacher"
    db.execute("INSERT INTO users(tg_id,name,role,phone) VALUES(?,?,?,?)",
               (m.from_user.id, m.from_user.full_name, role, m.contact.phone_number))
    db.commit()
    await m.answer(f"Ro'yxatdan muvaffaqiyatli o'tdingiz! ✅\nTelefon: {m.contact.phone_number}",
                   reply_markup=ReplyKeyboardRemove())
    await m.answer("Endi pastdagi 🔑 Kirish tugmasini bosing.", reply_markup=KIRISH_KB)

@dp.message(F.text == "🔑 Kirish")
async def kirish(m: Message):
    if not db.execute("SELECT 1 FROM users WHERE tg_id=?", (m.from_user.id,)).fetchone():
        await m.answer(greeting(m.from_user.first_name), reply_markup=CONTACT_KB()); return
    await menu(m.answer, role_of(m.from_user.id))

# ================= menyu =================
async def menu(answer, role, text="🏫 **DAVOMAT BOTI**\n"):
    buttons = [("✅ Davomat olish", "take"), ("📊 Hisobot", "report"),
               ("📈 Oylik foiz", "percent"), ("📤 Excel", "excel"),
               ("🗓 Dars jadvali", "lessons")]
    if role == "admin":
        buttons += [("➕ Guruh qo'shish", "add_group"),
                    ("👤 O'quvchi qo'shish", "add_student"),
                    ("🔐 O'qituvchini biriktirish", "assign"),
                    ("👥 Foydalanuvchini biriktirish", "assign_u"),
                    ("📚 Dars qo'shish", "add_lesson"),
                    ("📍 Bino qo'shish", "add_building"),
                    ("🔳 QR yaratish", "make_qr")]
    await answer(text, reply_markup=kb(buttons, 2), parse_mode="Markdown")

@dp.message(Command("menu"))
async def back(m: Message):
    if not db.execute("SELECT 1 FROM users WHERE tg_id=?", (m.from_user.id,)).fetchone():
        await m.answer(greeting(m.from_user.first_name), reply_markup=CONTACT_KB()); return
    await menu(m.answer, role_of(m.from_user.id))


# ================= guruh / o'quvchi =================
@dp.callback_query(F.data == "add_group")
async def add_group(c: CallbackQuery, state: FSMContext):
    await state.set_state(AddGroup.name)
    await c.message.answer("Yangi guruh (kurs/sinf) nomini yozing:")

@dp.message(AddGroup.name)
async def save_group(m: Message, state: FSMContext):
    db.execute("INSERT INTO groups_(name) VALUES(?)", (m.text.strip(),)); db.commit()
    await state.clear(); await m.answer(f"✅ Guruh «{m.text.strip()}» qo'shildi.")
    await menu(m.answer, role_of(m.from_user.id))

@dp.callback_query(F.data == "add_student")
async def add_student(c: CallbackQuery, state: FSMContext):
    if not db.execute("SELECT 1 FROM groups_").fetchone():
        return await c.message.answer("Avval guruh qo'shing!")
    await state.set_state(AddStudent.name)
    await c.message.answer("O'quvchining F.I.O ni yozing:")

@dp.message(AddStudent.name)
async def st_name(m: Message, state: FSMContext):
    await state.update_data(name=m.text.strip()); await state.set_state(AddStudent.group)
    rows = db.execute("SELECT id,name FROM groups_").fetchall()
    await m.answer("Qaysi guruhga?", reply_markup=kb([(n, f"stg_{i}") for i, n in rows]))

@dp.callback_query(AddStudent.group, F.data.startswith("stg_"))
async def st_group(c: CallbackQuery, state: FSMContext):
    await state.update_data(group=int(c.data.split("_")[1])); await state.set_state(AddStudent.parent)
    await c.message.answer("Ota-onasining Telegram ID sini yozing (/skip - o'tkazish):")

@dp.message(AddStudent.parent)
async def st_parent(m: Message, state: FSMContext):
    d = await state.get_data()
    pid = None if m.text == "/skip" else int(m.text)
    db.execute("INSERT INTO students(group_id,full_name,parent_id) VALUES(?,?,?)",
               (d["group"], d["name"], pid)); db.commit(); await state.clear()
    await m.answer(f"✅ O'quvchi «{d['name']}» qo'shildi.")
    await menu(m.answer, role_of(m.from_user.id))

# ================= biriktirish =================
@dp.callback_query(F.data == "assign")
async def assign(c: CallbackQuery, state: FSMContext):
    rows = db.execute("SELECT tg_id,name FROM users WHERE role='teacher'").fetchall()
    if not rows: return await c.message.answer("Hozircha o'qituvchilar yo'q.")
    await state.set_state(AssignT.teacher)
    await c.message.answer("O'qituvchini tanlang:", reply_markup=kb([(n, f"as_{i}") for i, n in rows]))

@dp.callback_query(AssignT.teacher, F.data.startswith("as_"))
async def assign2(c: CallbackQuery, state: FSMContext):
    await state.update_data(teacher=int(c.data.split("_")[1])); await state.set_state(AssignT.group)
    rows = db.execute("SELECT id,name FROM groups_").fetchall()
    await c.message.answer("Qaysi guruhga biriktirasiz?", reply_markup=kb([(n, f"ag_{i}") for i, n in rows]))

@dp.callback_query(AssignT.group, F.data.startswith("ag_"))
async def assign3(c: CallbackQuery, state: FSMContext):
    d = await state.get_data()
    db.execute("UPDATE users SET group_id=? WHERE tg_id=?", (int(c.data.split("_")[1]), d["teacher"]))
    db.commit(); await state.clear()
    await c.message.answer("✅ O'qituvchi biriktirildi.")
    await menu(c.message.answer, role_of(c.from_user.id))

@dp.callback_query(F.data == "assign_u")
async def assign_u(c: CallbackQuery, state: FSMContext):
    rows = db.execute("SELECT tg_id,name FROM users WHERE group_id IS NULL AND tg_id<>?",
                      (ADMIN_ID,)).fetchall()
    if not rows: return await c.message.answer("Barcha foydalanuvchilar guruhga biriktirilgan.")
    await state.set_state(AssignU.user)
    await c.message.answer("Foydalanuvchini tanlang:", reply_markup=kb([(n, f"au_{i}") for i, n in rows]))

@dp.callback_query(AssignU.user, F.data.startswith("au_"))
async def assign_u2(c: CallbackQuery, state: FSMContext):
    await state.update_data(user=int(c.data.split("_")[1])); await state.set_state(AssignU.group)
    rows = db.execute("SELECT id,name FROM groups_").fetchall()
    await c.message.answer("Qaysi guruhga?", reply_markup=kb([(n, f"agu_{i}") for i, n in rows]))

@dp.callback_query(AssignU.group, F.data.startswith("agu_"))
async def assign_u3(c: CallbackQuery, state: FSMContext):
    d = await state.get_data()
    db.execute("UPDATE users SET group_id=? WHERE tg_id=?", (int(c.data.split("_")[1]), d["user"]))
    db.commit(); await state.clear()
    await c.message.answer("✅ Foydalanuvchi guruhga biriktirildi.")
    await menu(c.message.answer, role_of(c.from_user.id))

# ================= bino (lokatsiya) =================
@dp.callback_query(F.data == "add_building")
async def add_building(c: CallbackQuery, state: FSMContext):
    await state.set_state(AddBuilding.name)
    await c.message.answer("Bino nomini yozing (masalan: A-korpus):")

@dp.message(AddBuilding.name)
async def b_name(m: Message, state: FSMContext):
    await state.update_data(name=m.text.strip()); await state.set_state(AddBuilding.lat)
    await m.answer("Binoning kengligini (latitude) yozing:\n"
                   "Google Maps da binoga bosib ko'rish mumkin. Masalan: 41.367259")

@dp.message(AddBuilding.lat)
async def b_lat(m: Message, state: FSMContext):
    try: float(m.text.strip())
    except ValueError: return await m.answer("❌ Raqam kiriting, masalan: 41.367259")
    await state.update_data(lat=float(m.text.strip())); await state.set_state(AddBuilding.lon)
    await m.answer("Binoning uzunligini (longitude) yozing, masalan: 69.396813")

@dp.message(AddBuilding.lon)
async def b_lon(m: Message, state: FSMContext):
    try: float(m.text.strip())
    except ValueError: return await m.answer("❌ Raqam kiriting, masalan: 69.396813")
    await state.update_data(lon=float(m.text.strip())); await state.set_state(AddBuilding.radius)
    await m.answer("Radiusni metrda yozing (masalan: 100 — bino atrofida 100 m):")

@dp.message(AddBuilding.radius)
async def b_radius(m: Message, state: FSMContext):
    try: r = int(m.text.strip())
    except ValueError: return await m.answer("❌ Butun son kiriting, masalan: 100")
    d = await state.get_data()
    db.execute("INSERT INTO buildings(name,lat,lon,radius_m) VALUES(?,?,?,?)",
               (d["name"], d["lat"], d["lon"], r)); db.commit(); await state.clear()
    await m.answer(f"✅ Bino qo'shildi: {d['name']}\n"
                   f"📍 {d['lat']}, {d['lon']} · radius: {r} m")
    await menu(m.answer, role_of(m.from_user.id))

def haversine_m(lat1, lon1, lat2, lon2):
    import math
    R = 6371000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1); dl = math.radians(lon2 - lon1)
    a = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2 * R * math.asin(math.sqrt(a))

# ================= QR yaratish (admin) =================
@dp.callback_query(F.data == "make_qr")
async def make_qr(c: CallbackQuery, state: FSMContext):
    rows = db.execute("SELECT id,name FROM groups_").fetchall()
    if not rows: return await c.message.answer("Avval guruh qo'shing!")
    await state.set_state(MakeQR.group)
    await c.message.answer("Qaysi guruh uchun QR?", reply_markup=kb([(n, f"qg_{i}") for i, n in rows]))

@dp.callback_query(MakeQR.group, F.data.startswith("qg_"))
async def make_qr2(c: CallbackQuery, state: FSMContext):
    gid = int(c.data.split("_")[1]); await state.update_data(group=gid)
    wd = datetime.datetime.now().weekday()
    rows = db.execute("""SELECT id,para_num,fan,start_time FROM lessons
                         WHERE group_id=? AND weekday=? ORDER BY start_time""", (gid, wd)).fetchall()
    if not rows: return await c.message.answer("Bugun bu guruhda dars yo'q!")
    await state.set_state(MakeQR.lesson)
    await c.message.answer("Qaysi para uchun?",
        reply_markup=kb([(f"{p}-para {f} ({t})", f"ql_{i}") for i, p, f, t in rows]))

@dp.callback_query(MakeQR.lesson, F.data.startswith("ql_"))
async def make_qr3(c: CallbackQuery, state: FSMContext):
    lid = int(c.data.split("_")[1]); d = await state.get_data(); await state.clear()
    lesson = db.execute("SELECT group_id,para_num FROM lessons WHERE id=?", (lid,)).fetchone()
    today = datetime.date.today().isoformat()
    import json as _json
    qr_text = _json.dumps({"l": lid, "g": lesson[0], "p": lesson[1], "d": today})
    import qrcode, io
    img = qrcode.make(qr_text)
    buf = io.BytesIO(); img.save(buf, format="PNG"); buf.seek(0)
    from aiogram.types import BufferedInputFile
    await c.message.answer_photo(
        BufferedInputFile(buf.read(), filename="qr.png"),
        caption=f"🔳 {lesson[1]}-para QR kodi\n"
                f"📅 Sana: {today}\n\n"
                f"Bu QR ni xonaga osib qo'ying. Kursantlar «Kirish»/«Chiqish» da skanerlaydi.")
    await menu(c.message.answer, role_of(c.from_user.id))

# ================= dars jadvali =================
@dp.callback_query(F.data == "add_lesson")
async def add_lesson(c: CallbackQuery, state: FSMContext):
    rows = db.execute("SELECT id,name FROM groups_").fetchall()
    if not rows: return await c.message.answer("Avval guruh qo'shing!")
    await state.set_state(AddLesson.group)
    await c.message.answer("Qaysi guruh uchun?", reply_markup=kb([(n, f"lg_{i}") for i, n in rows]))

@dp.callback_query(AddLesson.group, F.data.startswith("lg_"))
async def lg2(c: CallbackQuery, state: FSMContext):
    await state.update_data(group=int(c.data.split("_")[1])); await state.set_state(AddLesson.para)
    await c.message.answer("Para raqamini tanlang:", reply_markup=kb([(str(i), f"lp_{i}") for i in range(1, 7)], 3))

@dp.callback_query(AddLesson.para, F.data.startswith("lp_"))
async def lg3(c: CallbackQuery, state: FSMContext):
    await state.update_data(para=int(c.data.split("_")[1])); await state.set_state(AddLesson.fan)
    await c.message.answer("Fan nomini yozing:")

@dp.message(AddLesson.fan)
async def lg4(m: Message, state: FSMContext):
    await state.update_data(fan=m.text.strip()); await state.set_state(AddLesson.oqituvchi)
    await m.answer("O'qituvchi F.I.O:")

@dp.message(AddLesson.oqituvchi)
async def lg5(m: Message, state: FSMContext):
    await state.update_data(oqituvchi=m.text.strip()); await state.set_state(AddLesson.xona)
    await m.answer("Xona raqami:")

@dp.message(AddLesson.xona)
async def lg6(m: Message, state: FSMContext):
    await state.update_data(xona=m.text.strip()); await state.set_state(AddLesson.bino)
    await m.answer("Bino (masalan: A-korpus):")

@dp.message(AddLesson.bino)
async def lg7(m: Message, state: FSMContext):
    await state.update_data(bino=m.text.strip()); await state.set_state(AddLesson.weekday)
    await m.answer("Kunni tanlang:", reply_markup=kb([(WEEKDAYS[i], f"lw_{i}") for i in range(7)], 4))

@dp.callback_query(AddLesson.weekday, F.data.startswith("lw_"))
async def lg8(c: CallbackQuery, state: FSMContext):
    await state.update_data(weekday=int(c.data.split("_")[1])); await state.set_state(AddLesson.time)
    await c.message.answer("Boshlanish vaqtini HH:MM ko'rinishida yozing (masalan: 14:00):")

@dp.message(AddLesson.time)
async def lg9(m: Message, state: FSMContext):
    try: datetime.datetime.strptime(m.text.strip(), "%H:%M")
    except ValueError: return await m.answer("❌ Noto'g'ri format. Masalan: 14:00")
    d = await state.get_data()
    db.execute("""INSERT INTO lessons(group_id,para_num,fan,oqituvchi,xona,bino,weekday,start_time)
                  VALUES(?,?,?,?,?,?,?,?)""",
               (d["group"], d["para"], d["fan"], d["oqituvchi"], d["xona"], d["bino"],
                d["weekday"], m.text.strip()))
    db.commit(); await state.clear()
    await m.answer(f"✅ {d['para']}-para qo'shildi: {d['fan']}, {m.text.strip()}")
    await menu(m.answer, role_of(m.from_user.id))

@dp.callback_query(F.data == "lessons")
async def lessons(c: CallbackQuery):
    gid = group_of(c.from_user.id)
    if role_of(c.from_user.id) != "admin" and gid:
        return await show_lessons(c.message, gid)
    rows = db.execute("SELECT id,name FROM groups_").fetchall()
    if not rows: return await c.message.answer("Guruhlar yo'q.")
    await c.message.answer("Guruhni tanlang:", reply_markup=kb([(n, f"sv_{i}") for i, n in rows]))

@dp.callback_query(F.data.startswith("sv_"))
async def lessons2(c: CallbackQuery):
    await show_lessons(c.message, int(c.data.split("_")[1]))

async def show_lessons(message, gid):
    gname = db.execute("SELECT name FROM groups_ WHERE id=?", (gid,)).fetchone()[0]
    rows = db.execute("""SELECT para_num,fan,oqituvchi,xona,bino,weekday,start_time
                         FROM lessons WHERE group_id=? ORDER BY weekday, start_time""", (gid,)).fetchall()
    if not rows: return await message.answer("Bu guruhda darslar qo'yilmagan.")
    msg = f"🗓 **{gname} - dars jadvali**\n"
    for pn, fan, oqt, xona, bino, wd, t in rows:
        msg += f"\n{WEEKDAYS[wd]} {t} — **{pn}-para**\n📖 {fan}\n🏆 {oqt}\n🚪 {xona}, {bino}"
    await message.answer(msg, parse_mode="Markdown")

# ================= para boshlanishida avtomatik xabar =================
async def scheduler():
    while True:
        try:
            now = datetime.datetime.now()
            rows = db.execute("""SELECT l.group_id,l.para_num,l.fan,l.oqituvchi,l.xona,l.bino
                                 FROM lessons l WHERE l.weekday=? AND l.start_time=?""",
                              (now.weekday(), now.strftime("%H:%M"))).fetchall()
            for gid, pn, fan, oqt, xona, bino in rows:
                txt = (f"📚 **{pn}-para boshlanmoqda!**\n\n"
                       f"📖 Fan: {fan}\n🏆 O'qituvchi: {oqt}\n🚪 Xona: {xona}\n🏢 Bino: {bino}")
                for (tid,) in db.execute("SELECT tg_id FROM users WHERE group_id=?", (gid,)).fetchall():
                    try: await bot.send_message(tid, txt, parse_mode="Markdown")
                    except Exception: pass
        except Exception as e:
            logging.error(f"scheduler: {e}")
        await asyncio.sleep(30)

# ================= davomat =================
@dp.callback_query(F.data == "take")
async def take(c: CallbackQuery, state: FSMContext):
    gid = group_of(c.from_user.id)
    if role_of(c.from_user.id) != "admin" and gid:
        await state.update_data(group_id=gid); return await ask_date(c.message, state)
    rows = db.execute("SELECT id,name FROM groups_").fetchall()
    if not rows: return await c.message.answer("Guruhlar yo'q.")
    await c.message.answer("Guruhni tanlang:", reply_markup=kb([(n, f"t_{i}") for i, n in rows]))

@dp.callback_query(F.data.startswith("t_"))
async def take_group(c: CallbackQuery, state: FSMContext):
    await state.update_data(group_id=int(c.data.split("_")[1])); await ask_date(c.message, state)

async def ask_date(message, state):
    await state.set_state(TakeDate.date)
    today = datetime.date.today().isoformat()
    await message.answer("Sanani tanlang:", reply_markup=kb([("📅 Bugun", f"d_{today}"), ("✏️ Boshqa sana", "d_other")]))

@dp.callback_query(TakeDate.date, F.data.startswith("d_"))
async def set_date(c: CallbackQuery, state: FSMContext):
    if c.data == "d_other":
        await c.message.answer("Sanani YYYY-MM-DD ko'rinishida yozing:"); return
    await state.update_data(date=c.data.split("_")[1]); await send_students(c.message, state)

@dp.message(TakeDate.date)
async def set_date_text(m: Message, state: FSMContext):
    try: datetime.date.fromisoformat(m.text.strip())
    except ValueError: return await m.answer("❌ Noto'g'ri format. Masalan: 2026-09-23")
    await state.update_data(date=m.text.strip()); await send_students(m, state)

async def send_students(message, state):
    d = await state.get_data()
    students = db.execute("SELECT id,full_name FROM students WHERE group_id=?", (d["group_id"],)).fetchall()
    if not students: return await message.answer("Bu guruhda o'quvchi yo'q.")
    for sid, name in students:
        await message.answer(f"*{name}*", parse_mode="Markdown",
            reply_markup=kb([("✅ Keldi", f"a_{sid}_{d['date']}_keldi"),
                             ("❌ Kelmadi", f"a_{sid}_{d['date']}_kelmadi"),
                             ("📝 Sababli", f"a_{sid}_{d['date']}_sababli")], 3))
    await message.answer("Barcha o'quvchilar bo'yicha belgilab chiqing."); await state.clear()

@dp.callback_query(F.data.startswith("a_"))
async def mark(c: CallbackQuery):
    _, sid, date, status = c.data.split("_")
    db.execute("INSERT OR REPLACE INTO attendance(student_id,date,status) VALUES(?,?,?)",
               (int(sid), date, status)); db.commit()
    await c.message.edit_text(c.message.text + f"  →  *{status}*", parse_mode="Markdown")
    if status in ("kelmadi", "sababli"):
        pid = db.execute("SELECT parent_id FROM students WHERE id=?", (int(sid),)).fetchone()
        if pid and pid[0]:
            try:
                await bot.send_message(pid[0], "🔔 **Davomat xabari**\n\n"
                    f"Sizning farzandingiz {date} kuni darsga "
                    f"{'kelmadi ❌' if status=='kelmadi' else 'sababli kelmadi 📝'}.",
                    parse_mode="Markdown")
            except Exception: pass

# ================= hisobot / foiz / excel =================
@dp.callback_query(F.data == "report")
async def report(c: CallbackQuery):
    rows = db.execute("""SELECT g.name, s.full_name, a.status, COUNT(*)
        FROM attendance a JOIN students s ON s.id=a.student_id
        JOIN groups_ g ON g.id=s.group_id
        GROUP BY g.name, s.full_name, a.status""").fetchall()
    if not rows: return await c.message.answer("Hozircha ma'lumot yo'q.")
    data = {}
    for g, name, status, cnt in rows:
        data.setdefault((g, name), {})[status] = cnt
    msg = "📊 **HISOBOT**\n"
    for (g, name), st in sorted(data.items()):
        msg += f"\n*{g}* — {name}: ✅{st.get('keldi',0)} ❌{st.get('kelmadi',0)} 📝{st.get('sababli',0)}"
    await c.message.answer(msg, parse_mode="Markdown")

@dp.callback_query(F.data == "percent")
async def percent(c: CallbackQuery, state: FSMContext):
    rows = db.execute("SELECT id,name FROM groups_").fetchall()
    if not rows: return await c.message.answer("Guruhlar yo'q.")
    await state.set_state(PercentSt.group_id)
    await c.message.answer("Guruhni tanlang:", reply_markup=kb([(n, f"p_{i}") for i, n in rows]))

@dp.callback_query(PercentSt.group_id, F.data.startswith("p_"))
async def percent2(c: CallbackQuery, state: FSMContext):
    await state.update_data(group_id=int(c.data.split("_")[1])); await state.set_state(PercentSt.month)
    now = datetime.date.today(); months = []
    for i in range(3):
        y, mo = now.year, now.month - i
        while mo <= 0: mo += 12; y -= 1
        months.append((f"{y}-{mo:02d}", f"m_{y}-{mo:02d}"))
    await c.message.answer("Oyni tanlang:", reply_markup=kb(months))

@dp.callback_query(PercentSt.month, F.data.startswith("m_"))
async def percent3(c: CallbackQuery, state: FSMContext):
    month = c.data.split("_")[1]; gid = (await state.get_data())["group_id"]; await state.clear()
    gname = db.execute("SELECT name FROM groups_ WHERE id=?", (gid,)).fetchone()[0]
    students = db.execute("SELECT id,full_name FROM students WHERE group_id=?", (gid,)).fetchall()
    if not students: return await c.message.answer("Bu guruhda o'quvchi yo'q.")
    msg = f"📈 **{gname} — {month} oyi davomat foizi**\n"
    for sid, name in students:
        st = dict(db.execute("SELECT status,COUNT(*) FROM attendance "
                             "WHERE student_id=? AND date LIKE ? GROUP BY status",
                             (sid, month + "%")).fetchall())
        total = sum(st.values()); p = round(st.get("keldi",0)/total*100) if total else 0
        msg += f"\n{name}: **{p}%** ✅{st.get('keldi',0)} ❌{st.get('kelmadi',0)} 📝{st.get('sababli',0)}"
    await c.message.answer(msg, parse_mode="Markdown")

@dp.callback_query(F.data == "excel")
async def excel(c: CallbackQuery):
    rows = db.execute("""SELECT a.date, g.name, s.full_name, a.status
        FROM attendance a JOIN students s ON s.id=a.student_id
        JOIN groups_ g ON g.id=s.group_id ORDER BY a.date""").fetchall()
    if not rows: return await c.message.answer("Hozircha ma'lumot yo'q.")
    wb = Workbook(); ws = wb.active; ws.title = "Davomat"
    ws.append(["Sana", "Guruh", "O'quvchi", "Holat"])
    for r in rows: ws.append(list(r))
    fname = f"davomat_{datetime.date.today()}.xlsx"; wb.save(fname)
    await c.message.answer_document(FSInputFile(fname), caption="📤 Davomat hisoboti (Excel)")

INDEX_PATH = os.path.join(BASE, "index.html")
if not os.path.exists(INDEX_PATH):
    with open(INDEX_PATH, "w", encoding="utf-8") as f:
        f.write(INDEX_HTML)
    logging.info("index.html avtomatik yaratildi")

# ===== WEBHOOK (Render server) yoki POLLING (lokal sinov) =====
async def on_startup(app_):
    if PUBLIC_URL:
        try:
            await bot.set_webhook(f"{PUBLIC_URL}/webhook", drop_pending_updates=True)
            logging.info(f"Webhook o'rnatildi: {PUBLIC_URL}/webhook")
        except Exception as e:
            logging.error(f"Webhook xato: {e}")

async def webhook_handler(request):
    from aiogram.types import Update
    data = await request.json()
    await dp.feed_webhook_update(bot, Update.model_validate(data))
    return web.Response()

if RUN_WEB and HAS_WEB:
    app.router.add_post("/webhook", webhook_handler)
    app.on_startup.append(on_startup)

async def main():
    asyncio.create_task(scheduler())
    if RUN_WEB and HAS_WEB:
        runner = web.AppRunner(app)
        await runner.setup()
        site = web.TCPSite(runner, "0.0.0.0", PORT)
        await site.start()
        logging.info(f"Server: 0.0.0.0:{PORT}")
        if PUBLIC_URL:
            logging.info("Webhook rejim: server 24/7 tayyor")
            await asyncio.Event().wait()      # serverda doimiy kutish
        else:
            logging.info("Mahalliy rejim: polling")
            await dp.start_polling(bot)
    else:
        await dp.start_polling(bot)

@dp.message()
async def not_registered(m: Message):
    if not db.execute("SELECT 1 FROM users WHERE tg_id=?", (m.from_user.id,)).fetchone():
        await m.answer(f"{m.from_user.first_name}, siz hali ro'yxatdan o'tmagansiz.\n"
                       f"Quyidagi tugmani bosib ro'yxatdan o'ting:", reply_markup=CONTACT_KB())
    else:
        await menu(m.answer, role_of(m.from_user.id))

if __name__ == "__main__":
    asyncio.run(main())
