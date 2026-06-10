import os, base64, json, subprocess, tempfile, urllib.request, urllib.error
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

GH_TOKEN = os.environ.get("GH_TOKEN", "")
GH_REPO  = os.environ.get("GH_REPO", "adiRr-cell/forum-zero")
IV       = int(os.environ.get("INTERVAL_MS", "10000"))

@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "WHS Converter"})

@app.route("/convert", methods=["POST"])
def convert():
    try:
        data = request.get_json(force=True)
        pptx_b64 = data.get("pptx")
        name     = data.get("name", "Apresentação")
        action   = data.get("action", "add")  # "add" or "replace_id"
        replace_id = data.get("replace_id", None)

        if not pptx_b64:
            return jsonify({"error": "pptx ausente"}), 400

        # Decode pptx
        pptx_bytes = base64.b64decode(pptx_b64)

        with tempfile.TemporaryDirectory() as tmpdir:
            pptx_path = os.path.join(tmpdir, "input.pptx")
            with open(pptx_path, "wb") as f:
                f.write(pptx_bytes)

            # Convert to PDF via LibreOffice
            result = subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf",
                 "--outdir", tmpdir, pptx_path],
                capture_output=True, text=True, timeout=120
            )
            if result.returncode != 0:
                return jsonify({"error": "Falha na conversão: " + result.stderr}), 500

            pdf_path = os.path.join(tmpdir, "input.pdf")
            if not os.path.exists(pdf_path):
                return jsonify({"error": "PDF não gerado"}), 500

            # Convert PDF pages to JPEG
            result2 = subprocess.run(
                ["pdftoppm", "-jpeg", "-r", "120", pdf_path,
                 os.path.join(tmpdir, "slide")],
                capture_output=True, timeout=120
            )
            if result2.returncode != 0:
                return jsonify({"error": "Falha pdftoppm: " + result2.stderr}), 500

            # Collect slide images
            import glob
            slide_files = sorted(glob.glob(os.path.join(tmpdir, "slide-*.jpg")))
            if not slide_files:
                slide_files = sorted(glob.glob(os.path.join(tmpdir, "slide-*")))

            if not slide_files:
                return jsonify({"error": "Nenhum slide gerado"}), 500

            slides_b64 = []
            for sf in slide_files:
                with open(sf, "rb") as f:
                    slides_b64.append(base64.b64encode(f.read()).decode())

        # Load current state from GitHub
        state = load_state()

        from datetime import datetime
        now = datetime.now().strftime("%d/%m/%Y %H:%M")
        pres_id = "p" + str(int(datetime.now().timestamp() * 1000))

        new_pres = {
            "id": pres_id,
            "name": name,
            "date": now,
            "slideCount": len(slides_b64),
            "thumb": slides_b64[0],
            "slides": slides_b64
        }

        if action == "replace_id" and replace_id:
            state = [p for p in state if p["id"] != replace_id]

        state.append(new_pres)

        # Save state + rebuild TV
        save_state(state)
        rebuild_tv(state)

        return jsonify({
            "ok": True,
            "id": pres_id,
            "name": name,
            "slideCount": len(slides_b64),
            "thumb": slides_b64[0]
        })

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/delete", methods=["POST"])
def delete():
    try:
        data = request.get_json(force=True)
        del_id = data.get("id")
        if not del_id:
            return jsonify({"error": "id ausente"}), 400

        state = load_state()
        state = [p for p in state if p["id"] != del_id]
        save_state(state)
        rebuild_tv(state)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/state", methods=["GET"])
def get_state():
    try:
        state = load_state()
        # Return without slides data (too heavy)
        light = [{"id": p["id"], "name": p["name"], "date": p["date"],
                  "slideCount": p["slideCount"], "thumb": p["thumb"]} for p in state]
        return jsonify({"presentations": light})
    except Exception as e:
        return jsonify({"presentations": [], "error": str(e)})


# ── GITHUB HELPERS ──
def gh_get(path):
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GH_REPO}/{path}",
        headers={"Authorization": f"token {GH_TOKEN}",
                 "Accept": "application/vnd.github.v3+json"}
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())

def gh_put(path, content_str, message="WHS Update"):
    content_b64 = base64.b64encode(content_str.encode()).decode()
    sha = None
    try:
        cur = gh_get(path)
        sha = cur.get("sha")
    except: pass
    body = json.dumps({"message": message, "content": content_b64,
                       **({"sha": sha} if sha else {})}).encode()
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GH_REPO}/contents/{path}",
        data=body, method="PUT",
        headers={"Authorization": f"token {GH_TOKEN}",
                 "Accept": "application/vnd.github.v3+json",
                 "Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())

def load_state():
    try:
        r = gh_get("contents/state.json")
        return json.loads(base64.b64decode(r["content"].replace("\n","")).decode())
    except:
        return []

def save_state(state):
    # Save without slides (just metadata + thumb)
    meta = [{"id": p["id"], "name": p["name"], "date": p["date"],
             "slideCount": p["slideCount"], "thumb": p["thumb"]} for p in state]
    gh_put("state.json", json.dumps({"presentations": meta}))

def rebuild_tv(state):
    all_slides = []
    for p in state:
        all_slides.extend(p.get("slides", []))
    html = build_tv_html(all_slides, state)
    gh_put("index.html", html)

def build_tv_html(slides_b64, pres_list):
    total = len(slides_b64)
    if total == 0:
        return """<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Fórum Zero WHS</title>
<style>*{margin:0;padding:0}body{background:#001a3a;display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;color:#fff;font-family:Arial,sans-serif;gap:16px}</style></head>
<body><div style="font-size:3rem">📺</div><h2>Fórum Zero – WHS</h2><p style="color:#667799">Nenhuma apresentação ativa.</p></body></html>"""

    slides_js = "[\n" + ",\n".join(f'  "data:image/jpeg;base64,{s}"' for s in slides_b64) + "\n]"
    title = " · ".join(p["name"] for p in pres_list) if pres_list else "Fórum Zero WHS"

    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>Fórum Zero – WHS</title>
<style>*{{margin:0;padding:0;box-sizing:border-box}}body{{background:#000;font-family:Arial,sans-serif;overflow:hidden;width:100vw;height:100vh;display:flex;align-items:center;justify-content:center;cursor:none}}body.su{{cursor:default}}#si{{width:100%;height:100%;object-fit:contain;display:block;user-select:none}}#ui{{position:fixed;inset:0;pointer-events:none;opacity:0;transition:opacity .3s;z-index:10}}body.su #ui{{opacity:1;pointer-events:auto}}.ar{{position:absolute;top:50%;transform:translateY(-50%);background:rgba(0,0,0,.55);border:2px solid rgba(255,255,255,.3);color:#fff;font-size:2.5rem;width:70px;height:70px;border-radius:50%;display:flex;align-items:center;justify-content:center;cursor:pointer}}.ar:hover{{background:rgba(0,100,200,.75)}}#bp{{left:24px}}#bn{{right:24px}}#tb{{position:absolute;top:0;left:0;right:0;height:56px;background:linear-gradient(to bottom,rgba(0,0,0,.7),transparent);display:flex;align-items:center;justify-content:space-between;padding:0 24px}}#ctr{{color:#fff;font-size:1.1rem;font-weight:bold}}#ttl{{color:#ddd;font-size:.85rem}}#bfs{{background:rgba(0,0,0,.5);border:1px solid rgba(255,255,255,.3);color:#fff;font-size:1rem;padding:4px 10px;border-radius:6px;cursor:pointer}}#bb{{position:absolute;bottom:0;left:0;right:0;height:72px;background:linear-gradient(to top,rgba(0,0,0,.8),transparent);display:flex;align-items:center;justify-content:center;gap:16px;padding:0 100px}}#ts{{display:flex;gap:6px;overflow-x:auto;max-width:70vw;scrollbar-width:none}}#ts::-webkit-scrollbar{{display:none}}.th{{flex-shrink:0;width:72px;height:40px;object-fit:cover;border-radius:4px;border:2px solid transparent;cursor:pointer;opacity:.6}}.th:hover{{opacity:1}}.th.active{{border-color:#4da6ff;opacity:1}}#bpl{{flex-shrink:0;background:rgba(0,120,255,.8);border:2px solid rgba(255,255,255,.4);color:#fff;font-size:1.4rem;width:48px;height:48px;border-radius:50%;display:flex;align-items:center;justify-content:center;cursor:pointer}}#pw{{position:fixed;bottom:0;left:0;right:0;height:4px;background:rgba(255,255,255,.15);z-index:20}}#pb{{height:100%;background:#4da6ff;width:0%;transition:width linear}}#bdg{{position:fixed;bottom:12px;right:18px;color:rgba(255,255,255,.35);font-size:.78rem;pointer-events:none;z-index:5}}#tst{{position:fixed;top:70px;left:50%;transform:translateX(-50%);background:rgba(0,0,0,.7);color:#fff;padding:8px 20px;border-radius:20px;font-size:.9rem;pointer-events:none;opacity:0;transition:opacity .4s;z-index:30}}#tst.show{{opacity:1}}#fsp{{position:fixed;inset:0;background:#001a3a;display:flex;flex-direction:column;align-items:center;justify-content:center;z-index:100;cursor:pointer;gap:20px}}#fsp h1{{color:#fff;font-size:2.2rem;font-weight:bold;text-align:center}}#fsp p{{color:#aac8ff;font-size:1rem;text-align:center;max-width:480px}}#fsb{{background:#0066cc;color:#fff;border:none;border-radius:12px;padding:18px 48px;font-size:1.4rem;cursor:pointer;font-weight:bold}}#fsb:hover{{background:#0088ff}}#fsp.hidden{{display:none}}</style></head>
<body>
<div id="fsp"><h1>📺 Fórum Zero – WHS</h1><p>{title}</p><button id="fsb">▶ Iniciar</button><p style="font-size:.82rem;color:#667799">{total} slides · Clique para tela cheia</p></div>
<div style="width:100vw;height:100vh;display:flex;align-items:center;justify-content:center"><img id="si" src="" alt="" draggable="false"></div>
<div id="ui"><div id="tb"><span id="ctr">1/{total}</span><span id="ttl">{title}</span><button id="bfs">⛶ Tela cheia</button></div><button class="ar" id="bp">&#8592;</button><button class="ar" id="bn">&#8594;</button><div id="bb"><button id="bpl">⏸</button><div id="ts"></div></div></div>
<div id="pw"><div id="pb"></div></div><div id="bdg">1/{total}</div><div id="tst"></div>
<script>
const S={slides_js},T=S.length,IV={IV};let C=0,P=true,tmr=null,ut=null;
const si=document.getElementById('si'),ctr=document.getElementById('ctr'),bdg=document.getElementById('bdg'),ts=document.getElementById('ts'),bpl=document.getElementById('bpl'),pb=document.getElementById('pb'),tst=document.getElementById('tst');
S.forEach(s=>{{const i=new Image();i.src=s;}});
function gt(n){{C=((n%T)+T)%T;si.src=S[C];const l=(C+1)+'/'+T;ctr.textContent=l;bdg.textContent=l;document.querySelectorAll('.th').forEach((t,i)=>t.classList.toggle('active',i===C));const a=ts.querySelectorAll('.th')[C];if(a)a.scrollIntoView({{behavior:'smooth',block:'nearest',inline:'center'}});if(P)sa();}}
function sa(){{clearInterval(tmr);pb.style.transition='none';pb.style.width='0%';requestAnimationFrame(()=>requestAnimationFrame(()=>{{pb.style.transition='width '+IV+'ms linear';pb.style.width='100%';}}));tmr=setInterval(()=>gt(C+1),IV);}}
function stp(){{clearInterval(tmr);pb.style.transition='none';pb.style.width='0%';}}
function setP(v){{P=v;bpl.textContent=v?'⏸':'▶';v?sa():stp();shT(v?'▶ Autoplay':'⏸ Pausado');}}
function shT(m){{tst.textContent=m;tst.classList.add('show');clearTimeout(window._t);window._t=setTimeout(()=>tst.classList.remove('show'),1800);}}
S.forEach((s,i)=>{{const t=document.createElement('img');t.src=s;t.className='th'+(i===0?' active':'');t.addEventListener('click',()=>gt(i));ts.appendChild(t);}});
gt(0);
function eFS(){{const e=document.documentElement;if(e.requestFullscreen)e.requestFullscreen();else if(e.webkitRequestFullscreen)e.webkitRequestFullscreen();}}
function tFS(){{if(!document.fullscreenElement&&!document.webkitFullscreenElement)eFS();else if(document.exitFullscreen)document.exitFullscreen();}}
document.getElementById('fsb').addEventListener('click',()=>{{eFS();document.getElementById('fsp').classList.add('hidden');setP(true);}});
document.getElementById('fsp').addEventListener('click',e=>{{if(e.target!==document.getElementById('fsb')){{eFS();document.getElementById('fsp').classList.add('hidden');setP(true);}}}});
document.getElementById('bfs').addEventListener('click',tFS);
document.getElementById('bp').addEventListener('click',()=>gt(C-1));
document.getElementById('bn').addEventListener('click',()=>gt(C+1));
bpl.addEventListener('click',()=>setP(!P));
function shUI(){{document.body.classList.add('su');clearTimeout(ut);ut=setTimeout(()=>document.body.classList.remove('su'),3500);}}
document.addEventListener('mousemove',shUI);document.addEventListener('touchstart',shUI);
document.addEventListener('keydown',e=>{{shUI();switch(e.key){{case'ArrowRight':case'ArrowDown':case'PageDown':e.preventDefault();gt(C+1);break;case'ArrowLeft':case'ArrowUp':case'PageUp':e.preventDefault();gt(C-1);break;case' ':e.preventDefault();setP(!P);break;case'p':case'P':setP(!P);break;case'f':case'F':tFS();break;}}}});
let tx=null;document.addEventListener('touchstart',e=>{{tx=e.touches[0].clientX;}});document.addEventListener('touchend',e=>{{if(tx===null)return;const dx=e.changedTouches[0].clientX-tx;if(Math.abs(dx)>50){{dx<0?gt(C+1):gt(C-1);}}tx=null;}});
<\/script></body></html>"""


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
