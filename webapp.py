"""Web app version of the test bot — Telegram Mini App for Students & Teachers.

Routes:
  GET  /                       — subject list + teacher link
  GET  /test/{stem}            — interactive quiz page with timer
  POST /api/submit             — receive answers, save attempt, return redirect
  GET  /result/{attempt_id}    — final score & mistake review page
  GET  /teacher                — Teacher AI Test Generator & Excel Export UI
  POST /api/ai-generate        — generate & save quiz from topic/file (admin-verified)
  GET  /api/export.xlsx        — download all results as Excel (.xlsx)
"""
from __future__ import annotations

import hashlib
import hmac
import io
import json
import logging
import random
import tempfile
from pathlib import Path
from typing import Optional
from urllib.parse import parse_qsl

from aiohttp import web
from openpyxl import Workbook

import database as db
from ai_generator import extract_text_from_file, generate_quiz_txt
from config import ADMIN_IDS, BOT_TOKEN, SUBJECTS_DIR
from parser import list_subjects, load_subject

log = logging.getLogger(__name__)

QUESTIONS_PER_TEST = 30
TEST_DURATION_MIN = 30

_active_tests: dict[str, dict] = {}


def verify_init_data(init_data: str) -> Optional[dict]:
    """Verify Telegram WebApp initData; return user dict on success."""
    if not init_data:
        return None
    try:
        parsed = dict(parse_qsl(init_data, strict_parsing=False))
    except Exception:
        return None
    received_hash = parsed.pop("hash", None)
    if not received_hash:
        return None
    data_check_string = "\n".join(f"{k}={v}" for k, v in sorted(parsed.items()))
    secret_key = hmac.new(b"WebAppData", BOT_TOKEN.encode(), hashlib.sha256).digest()
    computed = hmac.new(
        secret_key, data_check_string.encode(), hashlib.sha256
    ).hexdigest()
    if not hmac.compare_digest(computed, received_hash):
        return None
    user_json = parsed.get("user")
    if not user_json:
        return None
    try:
        return json.loads(user_json)
    except json.JSONDecodeError:
        return None


def _esc(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _ball_5(pct: float) -> int:
    if pct >= 90:
        return 5
    if pct >= 75:
        return 4
    if pct >= 60:
        return 3
    return 2


# ─── pages ───

async def handle_index(request: web.Request) -> web.Response:
    subjects = list_subjects()
    if not subjects:
        cards = '<div class="empty-box">Hozircha fanlar yo\'q. O\'qituvchi panelidan AI orqali test qo\'shing.</div>'
    else:
        items = []
        for name, path in subjects:
            try:
                count = len(load_subject(path))
            except Exception:
                count = 0
            q_count = min(count, QUESTIONS_PER_TEST)
            items.append(
                f'<a href="/test/{path.stem}" class="card-item">'
                f'<div><div class="card-title">📘 {_esc(name)}</div>'
                f'<div class="card-sub">{count} savoldan {q_count} tasi · {TEST_DURATION_MIN} daqiqa</div></div>'
                f'<span class="pill">Boshlash →</span>'
                f'</a>'
            )
        cards = "\n".join(items)
    return web.Response(text=INDEX_HTML.replace("{{cards}}", cards), content_type="text/html")


async def handle_teacher_page(request: web.Request) -> web.Response:
    return web.Response(text=TEACHER_HTML, content_type="text/html")


async def handle_test_page(request: web.Request) -> web.Response:
    stem = request.match_info["stem"]
    path = SUBJECTS_DIR / f"{stem}.txt"
    if not path.exists():
        return web.Response(text="Fan topilmadi", status=404)
    all_q = load_subject(path)
    if not all_q:
        return web.Response(text="Bu fanda savollar yo'q.", status=404)

    if len(all_q) <= QUESTIONS_PER_TEST:
        selected = list(all_q)
        random.shuffle(selected)
    else:
        selected = random.sample(all_q, QUESTIONS_PER_TEST)

    display_name = stem.replace("_", " ").replace("-", " ").title()
    test_id = hashlib.md5(f"{stem}:{random.random()}".encode()).hexdigest()[:12]

    q_meta = []
    q_blocks = []
    for idx, q in enumerate(selected, 1):
        opts = list(q.options)
        random.shuffle(opts)
        correct_idx = next((i for i, (letter, _) in enumerate(opts) if letter == q.correct), 0)
        q_meta.append({
            "n": idx,
            "text": q.text,
            "options": [t for _, t in opts],
            "correct_idx": correct_idx,
        })
        opt_html = []
        for i, (_, opt_text) in enumerate(opts):
            opt_html.append(
                f'<label class="opt-label">'
                f'<input type="radio" name="q{idx}" value="{i}" onchange="updateProgress()">'
                f'<span>{_esc(opt_text)}</span>'
                f'</label>'
            )
        q_blocks.append(
            f'<div class="q-card">'
            f'<div class="q-num">Savol {idx} / {len(selected)}</div>'
            f'<div class="q-text">{_esc(q.text)}</div>'
            f'<div class="opts">{"".join(opt_html)}</div>'
            f'</div>'
        )

    _active_tests[test_id] = {
        "stem": stem,
        "display_name": display_name,
        "questions": q_meta,
    }
    if len(_active_tests) > 200:
        oldest = next(iter(_active_tests))
        _active_tests.pop(oldest, None)

    html = (
        TEST_HTML.replace("{{title}}", _esc(display_name))
        .replace("{{test_id}}", test_id)
        .replace("{{total}}", str(len(selected)))
        .replace("{{duration_sec}}", str(TEST_DURATION_MIN * 60))
        .replace("{{questions}}", "\n".join(q_blocks))
    )
    return web.Response(text=html, content_type="text/html")


async def handle_submit(request: web.Request) -> web.Response:
    data = await request.json()
    test_id = data.get("test_id", "")
    answers = data.get("answers", {})
    init_data = data.get("init_data", "")

    test = _active_tests.pop(test_id, None)
    if not test:
        return web.json_response({"error": "Test muddati tugagan yoki topilmadi."}, status=404)

    tg_user = verify_init_data(init_data)
    user_id = int(tg_user["id"]) if tg_user and "id" in tg_user else 0
    if tg_user:
        db.upsert_user(
            user_id,
            tg_user.get("username", ""),
            tg_user.get("first_name", ""),
            tg_user.get("last_name", ""),
        )

    correct_count = 0
    mistakes = []
    for q in test["questions"]:
        key = f"q{q['n']}"
        chosen = answers.get(key)
        if chosen is not None and int(chosen) == q["correct_idx"]:
            correct_count += 1
        else:
            chosen_text = q["options"][int(chosen)] if chosen is not None else "Javob berilmadi"
            mistakes.append({
                "n": q["n"],
                "text": q["text"],
                "chosen": chosen_text,
                "correct": q["options"][q["correct_idx"]],
            })

    total = len(test["questions"])
    attempt_id = db.start_attempt(
        user_id=user_id,
        chat_id=user_id,
        chat_type="webapp",
        subject=test["display_name"],
        total=total,
    )
    db.finish_attempt(attempt_id, correct_count)

    return web.json_response({
        "attempt_id": attempt_id,
        "correct": correct_count,
        "total": total,
        "mistakes": mistakes,
        "subject": test["display_name"],
    })


async def handle_result(request: web.Request) -> web.Response:
    attempt_id = int(request.match_info["attempt_id"])
    with db.get_conn() as conn:
        row = conn.execute("SELECT * FROM attempts WHERE id = ?", (attempt_id,)).fetchone()
    if not row:
        return web.Response(text="Natija topilmadi", status=404)

    correct = row["correct"]
    total = row["total"] or 1
    pct = (correct / total) * 100
    ball = _ball_5(pct)
    html = (
        RESULT_HTML.replace("{{subject}}", _esc(row["subject"]))
        .replace("{{correct}}", str(correct))
        .replace("{{total}}", str(total))
        .replace("{{pct}}", f"{pct:.0f}")
        .replace("{{ball}}", str(ball))
    )
    return web.Response(text=html, content_type="text/html")


async def handle_ai_generate(request: web.Request) -> web.Response:
    reader = await request.multipart()
    init_data = ""
    subject_name = ""
    topic_text = ""
    num_questions = 10
    difficulty = "o'rta"
    uploaded_bytes = None
    uploaded_name = ""

    while True:
        part = await reader.next()
        if part is None:
            break
        if part.name == "init_data":
            init_data = await part.text()
        elif part.name == "subject_name":
            subject_name = (await part.text()).strip()
        elif part.name == "topic_text":
            topic_text = (await part.text()).strip()
        elif part.name == "num_questions":
            try:
                num_questions = max(3, min(30, int(await part.text())))
            except ValueError:
                num_questions = 10
        elif part.name == "difficulty":
            difficulty = (await part.text()).strip() or "o'rta"
        elif part.name == "file" and part.filename:
            uploaded_name = part.filename
            uploaded_bytes = await part.read(decode=False)

    tg_user = verify_init_data(init_data)
    if ADMIN_IDS:
        if not tg_user or int(tg_user.get("id", 0)) not in ADMIN_IDS:
            return web.json_response({"error": "Faqat o'qituvchi (admin) test yarata oladi."}, status=403)

    source_text = topic_text
    if uploaded_bytes and uploaded_name:
        suffix = Path(uploaded_name).suffix.lower() or ".txt"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(uploaded_bytes)
            tmp_path = Path(tmp.name)
        try:
            extracted = extract_text_from_file(tmp_path)
            if extracted:
                source_text = (extracted + "\n\n" + source_text).strip()
        finally:
            tmp_path.unlink(missing_ok=True)

    if not source_text:
        return web.json_response({"error": "Mavzu matni yoki PDF/DOCX fayl kiriting."}, status=400)

    try:
        generated_txt = await generate_quiz_txt(
            source_text=source_text,
            num_questions=num_questions,
            difficulty=difficulty,
        )
    except Exception as e:
        return web.json_response({"error": str(e)}, status=500)

    raw_stem = subject_name or (Path(uploaded_name).stem if uploaded_name else "ai_test")
    safe_stem = "".join(c if c.isalnum() or c in "_-" else "_" for c in raw_stem.lower()).strip("_") or "ai_test"
    dest = SUBJECTS_DIR / f"{safe_stem}.txt"
    if dest.exists():
        dest = SUBJECTS_DIR / f"{safe_stem}_{random.randint(100, 999)}.txt"

    dest.write_text(generated_txt, encoding="utf-8")
    qs = load_subject(dest)
    if not qs:
        dest.unlink(missing_ok=True)
        return web.json_response({"error": "AI formatida savol ajratib bo'lmadi, qayta urinib ko'ring."}, status=500)

    return web.json_response({
        "ok": True,
        "filename": dest.name,
        "stem": dest.stem,
        "count": len(qs),
    })


async def handle_export_xlsx(request: web.Request) -> web.Response:
    wb = Workbook()
    ws = wb.active
    ws.title = "Natijalar"
    ws.append(["ID", "User ID", "Ism", "Username", "Fan", "To'g'ri", "Jami", "Foiz (%)", "Baho (5)", "Sana"])

    with db.get_conn() as conn:
        rows = conn.execute(
            """
            SELECT a.id, a.user_id, u.first_name, u.last_name, u.username,
                   a.subject, a.correct, a.total, a.finished_at
            FROM attempts a
            LEFT JOIN users u ON u.user_id = a.user_id
            WHERE a.finished_at IS NOT NULL
            ORDER BY a.finished_at DESC
            """
        ).fetchall()

    for r in rows:
        total = r["total"] or 1
        pct = round((r["correct"] / total) * 100, 1)
        full_name = " ".join(filter(None, [r["first_name"] or "", r["last_name"] or ""])).strip() or "WebApp User"
        uname = f"@{r['username']}" if r["username"] else ""
        ws.append([
            r["id"],
            r["user_id"],
            full_name,
            uname,
            r["subject"],
            r["correct"],
            r["total"],
            pct,
            _ball_5(pct),
            (r["finished_at"] or "")[:19].replace("T", " "),
        ])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return web.Response(
        body=buf.read(),
        headers={
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition": 'attachment; filename="test_natijalari.xlsx"',
        },
    )


# ─── HTML Templates ───

_BASE_CSS = """
:root {
  --bg: var(--tg-theme-bg-color, #0f172a);
  --card: var(--tg-theme-secondary-bg-color, #1e293b);
  --text: var(--tg-theme-text-color, #f8fafc);
  --hint: var(--tg-theme-hint-color, #94a3b8);
  --accent: var(--tg-theme-button-color, #3b82f6);
  --accent-text: var(--tg-theme-button-text-color, #ffffff);
}
* { box-sizing: border-box; }
body {
  margin: 0; padding: 16px;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  background: var(--bg); color: var(--text);
  line-height: 1.45;
}
.container { max-width: 680px; margin: 0 auto; }
.header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; gap: 10px; }
h1, h2 { margin: 0; font-size: 20px; }
.btn {
  display: inline-block; padding: 10px 14px; border-radius: 12px; border: none;
  background: var(--accent); color: var(--accent-text); font-weight: 600;
  text-decoration: none; cursor: pointer; font-size: 14px; text-align: center;
}
.btn-outline { background: transparent; border: 1px solid var(--accent); color: var(--accent); }
.card-item {
  display: flex; justify-content: space-between; align-items: center;
  background: var(--card); padding: 14px 16px; border-radius: 14px;
  margin-bottom: 10px; text-decoration: none; color: inherit;
}
.card-title { font-weight: 600; font-size: 16px; }
.card-sub { font-size: 12px; color: var(--hint); margin-top: 3px; }
.pill {
  background: var(--accent); color: var(--accent-text);
  padding: 5px 11px; border-radius: 20px; font-size: 12px; font-weight: 600;
}
.q-card { background: var(--card); padding: 16px; border-radius: 14px; margin-bottom: 12px; }
.q-num { font-size: 12px; color: var(--hint); font-weight: 600; margin-bottom: 6px; }
.q-text { font-size: 15px; font-weight: 600; margin-bottom: 12px; white-space: pre-wrap; }
.opt-label {
  display: flex; align-items: center; gap: 10px; padding: 10px 12px;
  border-radius: 10px; margin-bottom: 6px; cursor: pointer;
  background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08);
}
.opt-label input { accent-color: var(--accent); transform: scale(1.15); }
.sticky-top {
  position: sticky; top: 0; background: var(--bg); padding: 10px 0;
  z-index: 10; border-bottom: 1px solid rgba(255,255,255,0.08); margin-bottom: 14px;
}
.progress-bar { height: 6px; background: rgba(255,255,255,0.1); border-radius: 99px; overflow: hidden; margin-top: 8px; }
.progress-fill { height: 100%; background: var(--accent); width: 0%; transition: width 0.2s; }
.field { width: 100%; padding: 10px 12px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.15); background: var(--bg); color: var(--text); margin-bottom: 10px; font-size: 14px; }
.empty-box { background: var(--card); padding: 20px; border-radius: 14px; text-align: center; color: var(--hint); }
"""

INDEX_HTML = f"""<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Maktab Test Mini App</title>
<script src="https://telegram.org/js/telegram-web-app.js"></script>
<style>{_BASE_CSS}</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>📚 Fanlar va Testlar</h1>
    <a href="/teacher" class="btn btn-outline">👩‍🏫 O'qituvchi</a>
  </div>
  {{cards}}
</div>
<script>if (window.Telegram && Telegram.WebApp) {{ Telegram.WebApp.ready(); Telegram.WebApp.expand(); }}</script>
</body>
</html>"""

TEST_HTML = f"""<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{title}}</title>
<script src="https://telegram.org/js/telegram-web-app.js"></script>
<style>{_BASE_CSS}</style>
</head>
<body>
<div class="container">
  <div class="sticky-top">
    <div style="display:flex;justify-content:space-between;align-items:center;">
      <strong>📘 {{title}}</strong>
      <span id="timer" style="font-weight:700;color:var(--accent);">--:--</span>
    </div>
    <div class="card-sub" id="progText">0 / {{total}} javob belgilandi</div>
    <div class="progress-bar"><div class="progress-fill" id="progFill"></div></div>
  </div>
  <form id="quizForm">
    {{questions}}
    <button type="submit" class="btn" style="width:100%;padding:14px;font-size:16px;margin-top:8px;">✅ Testni yakunlash</button>
  </form>
</div>
<script>
const tg = window.Telegram && Telegram.WebApp ? Telegram.WebApp : null;
if (tg) {{ tg.ready(); tg.expand(); }}
const TOTAL = {{total}};
let sec = {{duration_sec}};
function tick() {{
  const m = String(Math.floor(sec / 60)).padStart(2, '0');
  const s = String(sec % 60).padStart(2, '0');
  document.getElementById('timer').textContent = m + ':' + s;
  if (sec <= 0) {{ submitQuiz(); return; }}
  sec--;
  setTimeout(tick, 1000);
}}
tick();
function updateProgress() {{
  let answered = 0;
  for (let i = 1; i <= TOTAL; i++) {{
    if (document.querySelector('input[name="q' + i + '"]:checked')) answered++;
  }}
  document.getElementById('progText').textContent = answered + ' / ' + TOTAL + ' javob belgilandi';
  document.getElementById('progFill').style.width = (answered / TOTAL * 100) + '%';
}}
async function submitQuiz() {{
  const answers = {{}};
  for (let i = 1; i <= TOTAL; i++) {{
    const el = document.querySelector('input[name="q' + i + '"]:checked');
    if (el) answers['q' + i] = parseInt(el.value, 10);
  }}
  const res = await fetch('/api/submit', {{
    method: 'POST',
    headers: {{'Content-Type': 'application/json'}},
    body: JSON.stringify({{ test_id: '{{test_id}}', answers, init_data: tg ? tg.initData : '' }})
  }});
  const data = await res.json();
  if (data.attempt_id) {{
    sessionStorage.setItem('mistakes_' + data.attempt_id, JSON.stringify(data.mistakes || []));
    window.location.href = '/result/' + data.attempt_id;
  }} else {{
    alert(data.error || 'Xatolik yuz berdi');
  }}
}}
document.getElementById('quizForm').addEventListener('submit', (e) => {{ e.preventDefault(); submitQuiz(); }});
</script>
</body>
</html>"""

RESULT_HTML = f"""<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Natija</title>
<script src="https://telegram.org/js/telegram-web-app.js"></script>
<style>{_BASE_CSS}</style>
</head>
<body>
<div class="container">
  <div class="q-card" style="text-align:center;">
    <h2>🏁 {{subject}}</h2>
    <div style="font-size:42px;font-weight:800;margin:12px 0;color:var(--accent);">{{pct}}%</div>
    <p>To'g'ri javoblar: <strong>{{correct}} / {{total}}</strong> · Baho: <strong>{{ball}}</strong></p>
    <a href="/" class="btn" style="margin-top:8px;">📚 Boshqa testlar</a>
  </div>
  <div id="mistakesBox"></div>
</div>
<script>
if (window.Telegram && Telegram.WebApp) Telegram.WebApp.ready();
const id = window.location.pathname.split('/').pop();
const mistakes = JSON.parse(sessionStorage.getItem('mistakes_' + id) || '[]');
if (mistakes.length) {{
  let html = '<h2 style="margin:16px 0 10px;">❌ Xato qilingan savollar</h2>';
  for (const m of mistakes) {{
    html += `<div class="q-card"><div class="q-num">Savol ${{m.n}}</div><div class="q-text">${{m.text}}</div>` +
            `<div style="color:#f87171;font-size:13px;">Sizning javob: ${{m.chosen}}</div>` +
            `<div style="color:#4ade80;font-size:13px;margin-top:4px;">To'g'ri javob: ${{m.correct}}</div></div>`;
  }}
  document.getElementById('mistakesBox').innerHTML = html;
}}
</script>
</body>
</html>"""

TEACHER_HTML = f"""<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>O'qituvchi Paneli</title>
<script src="https://telegram.org/js/telegram-web-app.js"></script>
<style>{_BASE_CSS}</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>👩‍🏫 O'qituvchi paneli</h1>
    <a href="/" class="btn btn-outline">⬅️ Testlar</a>
  </div>

  <div class="q-card">
    <h2 style="margin-bottom:10px;">🤖 AI Test Generator</h2>
    <p class="card-sub" style="margin-bottom:12px;">PDF, Word (.docx) fayl yuklang yoki mavzu matnini yozing — Gemini AI avtomatik test tuzadi.</p>
    <form id="aiForm">
      <input class="field" name="subject_name" placeholder="Fan / mavzu nomi (masalan: 8-sinf Fizika)" required>
      <div style="display:flex;gap:8px;">
        <select class="field" name="num_questions">
          <option value="5">5 ta savol</option>
          <option value="10" selected>10 ta savol</option>
          <option value="15">15 ta savol</option>
          <option value="20">20 ta savol</option>
        </select>
        <select class="field" name="difficulty">
          <option value="oson">Oson</option>
          <option value="o'rta" selected>O'rta</option>
          <option value="qiyin">Qiyin</option>
        </select>
      </div>
      <input class="field" type="file" name="file" accept=".pdf,.docx,.txt">
      <textarea class="field" name="topic_text" rows="4" placeholder="Yoki mavzu matnini / tayanch konspektni shu yerga yozing..."></textarea>
      <button type="submit" class="btn" id="genBtn" style="width:100%;">✨ AI bilan test yaratish</button>
    </form>
    <div id="aiStatus" style="margin-top:10px;font-size:14px;"></div>
  </div>

  <div class="q-card">
    <h2 style="margin-bottom:8px;">📊 Natijalar hisoboti</h2>
    <p class="card-sub" style="margin-bottom:12px;">Barcha o'quvchilarning test natijalarini Excel (.xlsx) formatda yuklab oling.</p>
    <a href="/api/export.xlsx" class="btn" style="width:100%;">📥 Excel (.xlsx) yuklab olish</a>
  </div>
</div>
<script>
const tg = window.Telegram && Telegram.WebApp ? Telegram.WebApp : null;
if (tg) {{ tg.ready(); tg.expand(); }}
document.getElementById('aiForm').addEventListener('submit', async (e) => {{
  e.preventDefault();
  const btn = document.getElementById('genBtn');
  const status = document.getElementById('aiStatus');
  btn.disabled = true;
  btn.textContent = '⏳ AI test tuzmoqda...';
  status.textContent = '';
  const fd = new FormData(e.target);
  fd.append('init_data', tg ? tg.initData : '');
  try {{
    const res = await fetch('/api/ai-generate', {{ method: 'POST', body: fd }});
    const data = await res.json();
    if (data.ok) {{
      status.innerHTML = `<span style="color:#4ade80;">✅ Tayyor! <strong>${{data.filename}}</strong> (${{data.count}} ta savol) qo'shildi.</span>`;
    }} else {{
      status.innerHTML = `<span style="color:#f87171;">❌ ${{data.error}}</span>`;
    }}
  }} catch (err) {{
    status.innerHTML = `<span style="color:#f87171;">❌ Tarmoq xatosi: ${{err}}</span>`;
  }} finally {{
    btn.disabled = false;
    btn.textContent = '✨ AI bilan test yaratish';
  }}
}});
</script>
</body>
</html>"""


async def start_webapp(host: str = "0.0.0.0", port: int = 8080) -> web.AppRunner:
    app = web.Application(client_max_size=15 * 1024 * 1024)
    app.router.add_get("/", handle_index)
    app.router.add_get("/teacher", handle_teacher_page)
    app.router.add_get("/test/{stem}", handle_test_page)
    app.router.add_post("/api/submit", handle_submit)
    app.router.add_get("/result/{attempt_id}", handle_result)
    app.router.add_post("/api/ai-generate", handle_ai_generate)
    app.router.add_get("/api/export.xlsx", handle_export_xlsx)

    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, host, port)
    await site.start()
    log.info(f"WebApp http://{host}:{port} da ishga tushdi")
    return runner
