import sys, os, io, math, re, json, time, uuid, shutil, tempfile, threading, subprocess
import importlib, importlib.util, zipfile, html as _html, traceback
from pathlib import Path
from html.parser import HTMLParser

APP = "Convi"
VERSION = "3.2"
IS_WIN = os.name == "nt"
FROZEN = getattr(sys, "frozen", False)
NOWIN = 0x08000000 if IS_WIN else 0

if sys.stdout is None: sys.stdout = open(os.devnull, "w")
if sys.stderr is None: sys.stderr = open(os.devnull, "w")

REQUIRED = [("PIL", "Pillow"), ("imageio_ffmpeg", "imageio-ffmpeg"), ("docx", "python-docx"),
            ("reportlab", "reportlab"), ("pypdf", "pypdf"), ("markdown", "Markdown"),
            ("send2trash", "Send2Trash"), ("PySide6", "PySide6"), ("openpyxl", "openpyxl"), ("pptx", "python-pptx")]
OPTIONAL = [("pillow_heif", "pillow-heif"), ("xlrd", "xlrd"), ("fitz", "PyMuPDF")]
if IS_WIN: OPTIONAL.append(("win32com", "pywin32"))

def bootstrap(verbose=False):
    if FROZEN: return True
    todo = [(m, p, True) for m, p in REQUIRED if importlib.util.find_spec(m) is None]
    todo += [(m, p, False) for m, p in OPTIONAL if importlib.util.find_spec(m) is None]
    ok = True
    for m, p, req in todo:
        if verbose: print(f"  установка: {p} ...", flush=True)
        cmd = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "-q", p]
        if sys.prefix == getattr(sys, "base_prefix", sys.prefix): cmd.insert(4, "--user")
        r = subprocess.run(cmd, capture_output=not verbose, creationflags=NOWIN)
        if r.returncode != 0 and req:
            ok = False
            if verbose: print(f"  ОШИБКА: не удалось установить {p}", flush=True)
    importlib.invalidate_caches()
    return ok

DATA = Path(os.environ.get("APPDATA") or (Path.home() / ".config")) / "Convi"
DEFAULTS = dict(mode="near", trash=True, autostart=False, ctx=True, notify=True, office=True, hk_single="Ctrl+F+H", hk_batch="Ctrl+F+V", outdir_path="", quality=90, dpi=150, merge_pdf=True, last={}, watch_on=False, watch_dir="", watch_img="JPG", watch_vid="", watch_aud="", watch_doc="")
SET_LOCK = threading.Lock()
S = dict(DEFAULTS)

def load_settings():
    DATA.mkdir(parents=True, exist_ok=True)
    try: S.update({k: v for k, v in json.loads((DATA / "settings.json").read_text("utf-8")).items() if k in DEFAULTS})
    except Exception: pass

def save_settings():
    with SET_LOCK:
        try: (DATA / "settings.json").write_text(json.dumps(S, ensure_ascii=False, indent=1), "utf-8")
        except Exception: pass

def log(*a):
    try:
        DATA.mkdir(parents=True, exist_ok=True)
        with open(DATA / "convi.log", "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + " ".join(str(x) for x in a) + "\n")
    except Exception: pass

def install_error_hooks():
    def eh(t, v, tb): log("ОШИБКА:\n" + "".join(traceback.format_exception(t, v, tb)))
    sys.excepthook = eh
    try: threading.excepthook = lambda a: eh(a.exc_type, a.exc_value, a.exc_traceback)
    except Exception: pass
    try:
        lg = DATA / "convi.log"
        if lg.exists() and lg.stat().st_size > 300000: lg.write_text("", "utf-8")
    except Exception: pass

def default_outdir():
    for p in (Path.home() / "Downloads", Path.home() / "Documents", Path.home()):
        if p.exists(): return p / "Convi"
    return Path.home() / "Convi"

def _s(x): return set(x.split())
IMG_PIL = _s("""
apng avif avifs blp bmp bw cur dcx dds dib exr fit fits flc fli ftc ftu gbr gif hdr heic heif hif
    icb icns ico iim im j2c j2k jfi jfif jif jp2 jpc jpe jpeg jpf jpg jpx jxl mpo msp palm pam pbm
    pcd pcx pfm pgm pic pict png pnm ppm psd pxr qoi ras rgb rgba sgi sun tga tif tiff vda wbmp webp
    xbm xpm xwd""")
IMG_VEC = _s("""
ai cdr cgm cmx dxf emf eps fh fodg odg otg ps pub std svg svgz sxd vdx vsd vsdm vsdx wmf wpg xcf""")
VIDEO = _s("""
264 265 3g2 3gp 3gp2 3gpp 3gpp2 amv av1 avc avi bik cin dat divx drc dv dvr dvr-ms evo f4p f4v film
    flic flv flx gxf h264 h265 hevc isma ismv ivf ivr ktx lrf lrv m1v m2p m2t m2ts m2v m4v mjp mjpeg
    mjpg mk3d mkv mod mov mp2v mp4 mp4v mpe mpeg mpg mpg4 mpl mpv mpv2 mqv mts mtv mv mvi mxf nsv
    nut obu ogm ogv ogx pva qt rec rm rms rmvb roq rv sbk seq smk str thp tod tp trp   tvs usm vc1
    vc2 vdr vid vob vp6 vp8 vp9 webm wm wmv wmx wtv wvx xesc xmv xvid y4m yuv""")
AUDIO = _s("""
3ga aa aac aax ac3 act adt adts adx afc aif aifc aiff amr ape at3 at9 au avr awb bit brstm bwav bwf
    c2 caf cdda daud dff dsd dsf dss dtk dts dtshd dtsma dvd eac3 ec3 fap ffmetadata fla flac g722
    g723 g726 g729 gdm gsm hca ircam kar la lyra m4a m4b m4p m4r mac mid midi mka mks mlp mmf mp+
    mp1 mp2 mp3 mp4a mpa mpc mpga mpp msbc mt2 ofr ofs oga ogg oma omg opus paf pcm pvf qcp ra ram
    rax rf64 rsd rso sbc sd2 sds sf sln smp snd spdif sph spx tak tco thd truehd tta tvd txw vag voc
    vox vqf w64 wady wav wave weba wma wv wvc wve xa xmf xvag""")
DOC_B = _s("""
docx htm html log markdown md mdown mkd odt pdf rtf text txt xhtml""")
DOC_LO = _s("""
602 abw doc docm dot dotm dotx fb2 fodt hwp lwp mcw mw odm oth ott pages pdb psw sdw sgl stw sxg sxw
    uot vor wpd wps wpt zabw""")
TEXT = _s("""
abap ada adb adoc ads agda asc asciidoc ascx asd ashx asm asp aspx ass astro atom authors awk axaml
    babelrc bas bash bat bazel bbl bib bicep bsv build bzl c c++ cabal cairo cbl cc cfc cfg cfm
    changelog cjs cl clj cljc cljs cls cmake cmd cob coffee conf config coq cpp cr cs csh cshtml
    cson csproj css cue cxx d dart ddl dfm dhall diff diz dockerfile dockerignore dpr dtd dtx
    editorconfig edn ejs el elm entitlements env erb erl eslintrc ex exs f f03 f08 f90 f95 fasl fish
    for frm fs fsi fsproj fsscript fsx ftn fun gawk gemfile gemspec geojson gitattributes gitignore
    gitmodules gn gni go gp gql gradle graphql groovy h h++ haml hbs hcl hh hpp hrl hs htaccess hx
    hxml hxx iced idr inc ini ino ins ipynb jade java jenkinsfile jl js json json5 jsonc jsonl
    jsonnet jsp jspx jsx ksh kt kts latex lean less lhs libsonnet license liquid lisp litcoffee lock
    lpr lrc ls lsp ltx lua m make makefile master mdx meson mjs mk ml mli mll mlt mly mm move
    mustache nasm ndjson nfo nim ninja nix njk nl npmrc ocaml opam org owl p4 pas patch pbxproj pde
    php php3 php4 php5 phtml pl plg plist plsql pm pod podspec pp prettierrc prn pro procfile prolog
    properties props proto ps1 psd1 psm1 psql pug purs pxd py pyi pyw pyx r rake rakefile razor rb
    rbw rdf re readme reason rego rei rej res resx rkt rmd rs rss rst s sass sbt sbv sc scala scm
    scrbl scss sed sh slim smi sml sol sql srt ss ssa stack star strings sty styl sub sum sv svelte
    svh swift t tab targets tcl tcsh tex tf tfvars thrift thy tk tm todo toml tpl tsx twig v
    vagrantfile vb vba vbe vbproj vbs vcxproj verilog vhd vhdl vtt vue vy workspace wsf xaml
    xcconfig xml xsd xsl xslt yaml yml zig zsh""")
TAB_B = _s("""
csv ods tsv xls xlsm xlsx""")
TAB_LO = _s("""
123 dbf dif fods gnumeric numbers ots qpw slk stc sxc sylk uos wb1 wb2 wk1 wk3 wk4 wks wq1 wq2 xlsb
    xlt xltm xltx xlw""")
PRES_B = _s("""
pptx""")
PRES_LO = _s("""
dps dpt fodp key odp otp pot potm potx pps ppsm ppsx ppt pptm prz sda sdd shw sti sxi uop""")
PROBE = _s("""
266 302 722 aa3 ac4 acm adf adp aea aix al ans apc apl aptx aptxhd art asf ast avs avs2 avs3 bcstm
    bfstm binka bmv bonk cdata cdg cdxl cgi cif dav dfpwm evc flm fsb fwse g723_1 genh h261 h266
    h26l iamf ice idf ifv imx ipu ism kux laf lvf m2a mca mcc mj2 mods moflex mpl2 msf mtaf musx mxg
    nist nsp osq pdv pjs psp qcif qoa rco rcv rka rt sami sb sbg scc sdns sdr2 sdx ser sga shn son
    ss2 stl sup svag svs sw ty ty+ ub ul uw v210 viv vpk vqe vql vt vvc wa way wsd xl xmd yop yuv10""")

EXT_TYPE = {}
for _t, _g in (("image", IMG_PIL | IMG_VEC), ("video", VIDEO), ("audio", AUDIO), ("doc", DOC_B | DOC_LO | TEXT),
               ("table", TAB_B | TAB_LO), ("pres", PRES_B | PRES_LO)):
    for _e in _g: EXT_TYPE.setdefault(_e, _t)
ALL_EXT = set(EXT_TYPE) | PROBE
NEED_OFFICE = DOC_LO | TAB_LO | PRES_LO | (IMG_VEC - {"svg", "svgz"})
TYPES = {
 "image": dict(label="Фото", short="Фото", icon="image", to=["JPG","PNG","WEBP","AVIF","BMP","GIF","TIFF","ICO","PDF"]),
 "video": dict(label="Видео", short="Видео", icon="film", to=["MP4","AVI","MKV","MOV","WEBM","WMV","GIF","MP3","WAV"]),
 "audio": dict(label="Аудио", short="Аудио", icon="music", to=["MP3","WAV","OGG","FLAC","M4A","AAC","OPUS","WMA"]),
 "doc": dict(label="Документы", short="Док", icon="doc", to=["PDF","DOCX","TXT","HTML","MD","RTF","JPG","PNG"]),
 "table": dict(label="Таблицы", short="Табл", icon="table", to=["XLSX","CSV","PDF","HTML","JSON"]),
 "pres": dict(label="Презентации", short="Слайды", icon="pres", to=["PDF","PPTX","TXT","DOCX"]),
}
for _t, _v in TYPES.items(): _v["ext"] = sorted(e for e, x in EXT_TYPE.items() if x == _t)
ALIAS = {"jpeg": "jpg", "tif": "tiff", "htm": "html", "markdown": "md", "aif": "aiff", "text": "txt", "oga": "ogg", "dib": "bmp"}
def ext_of(n): return Path(n).suffix.lower().lstrip(".")
def type_of(n): return EXT_TYPE.get(ext_of(n))
def same_fmt(e, f): return ALIAS.get(e.lower(), e.lower()) == ALIAS.get(f.lower(), f.lower())
def outputs_for(name, t=None):
    t = t or type_of(name)
    if not t: return []
    out = list(TYPES[t]["to"])
    return [f for f in out if f not in ("JPG", "PNG")] if t == "doc" and ext_of(name) != "pdf" else out
def targets(name, t=None):
    e = ext_of(name); return [f for f in outputs_for(name, t) if not same_fmt(e, f)]

_SN = {}
def file_type(path):
    p = Path(path); t = EXT_TYPE.get(ext_of(p.name))
    if t: return t
    k = str(p)
    if k not in _SN: _SN[k] = _sniff(p)
    return _SN[k]

def _sniff(p):
    try:
        from PIL import Image
        with Image.open(p): return "image"
    except Exception: pass
    ff, _ = find_ffmpeg()
    if ff:
        try: out = subprocess.run([ff, "-hide_banner", "-i", str(p)], capture_output=True, text=True, encoding="utf-8", errors="replace", creationflags=NOWIN, timeout=20).stderr
        except Exception: out = ""
        if "Video:" in out or "Audio:" in out:
            return "video" if ("Video:" in out and "attached pic" not in out and not re.search(r"Video: (mjpeg|png|bmp)", out)) else "audio"
    try:
        if p.stat().st_size <= 30000000:
            head = p.read_bytes()[:8192]
            if head and b"\0" not in head: return "doc"
    except Exception: pass
    return None

def stat_path(p):
    p = Path(p)
    if not p.is_file(): return None
    return dict(path=str(p), name=p.name, size=p.stat().st_size, type=file_type(p))

_ENG_CACHE = {}
def find_ffmpeg():
    try:
        import imageio_ffmpeg
        p = imageio_ffmpeg.get_ffmpeg_exe()
        if p and os.path.exists(p): return p, "встроен"
    except Exception: pass
    p = shutil.which("ffmpeg")
    return (p, "системный") if p else (None, None)

def find_soffice():
    for n in ("soffice", "soffice.exe", "libreoffice"):
        p = shutil.which(n)
        if p: return p
    if IS_WIN:
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"), os.environ.get("LOCALAPPDATA")):
            if base:
                c = Path(base) / "LibreOffice" / "program" / "soffice.exe"
                if c.exists(): return str(c)
    return None

def word_ok():
    if not IS_WIN: return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"))
        return importlib.util.find_spec("win32com") is not None
    except Exception: return False

def has(mod):
    try:
        if importlib.util.find_spec(mod) is not None: return True
    except Exception: pass
    try: importlib.import_module(mod); return True
    except Exception: return False

def engines():
    ff, label = find_ffmpeg()
    return dict(ffmpeg=label if ff else None, pillow=has("PIL"), word=word_ok(), libre=bool(find_soffice()),
                docs=all(has(m) for m in ("docx", "reportlab", "pypdf", "markdown")), heif=has("pillow_heif"),
                tables=has("openpyxl"), xls=has("xlrd"), pres=has("pptx"), ppt=ppt_ok(), pdfimg=has("fitz"))

def ff_encoders(ff):
    if "enc" not in _ENG_CACHE:
        try:
            out = subprocess.run([ff, "-hide_banner", "-encoders"], capture_output=True, text=True,
                                 encoding="utf-8", errors="replace", creationflags=NOWIN).stdout
            _ENG_CACHE["enc"] = set(re.findall(r"^\s[VAS][\w.]{5}\s+(\S+)", out, re.M))
        except Exception: _ENG_CACHE["enc"] = set()
    return _ENG_CACHE["enc"]

class Cancelled(Exception): pass

class Job:
    def __init__(self, items):
        self.id = uuid.uuid4().hex[:10]; self.items = items; self.state = "run"
        self.pct = 0.0; self.idx = 0; self.total = len(items); self.cur = ""
        self.cancel = False; self.proc = None; self.results = []
    def check(self):
        if self.cancel: raise Cancelled()
    def kill(self):
        self.cancel = True
        try:
            if self.proc and self.proc.poll() is None: self.proc.terminate()
        except Exception: pass
    def view(self):
        return dict(state=self.state, pct=round(self.pct, 4), idx=self.idx, total=self.total, cur=self.cur, results=self.results)

JOBS = {}

def unique(p):
    p = Path(p)
    if not p.exists(): return p
    i = 1
    while True:
        q = p.with_name(f"{p.stem} ({i}){p.suffix}")
        if not q.exists(): return q
        i += 1

def short_err(e):
    s = str(e).strip() or e.__class__.__name__
    s = re.sub(r"\s+", " ", s)
    return s[:300]

def svg_to_image(path):
    from PySide6.QtSvg import QSvgRenderer
    from PySide6.QtGui import QImage, QPainter
    from PIL import Image
    r = QSvgRenderer(str(path))
    if not r.isValid(): raise RuntimeError("Не удалось прочитать SVG")
    sz = r.defaultSize(); w, h = sz.width() or 1024, sz.height() or 1024
    k = max(1.0, min(4.0, 1600 / max(w, h))); W, H = int(w * k), int(h * k)
    img = QImage(W, H, QImage.Format.Format_ARGB32); img.fill(0)
    pt = QPainter(img); r.render(pt); pt.end()
    return Image.frombuffer("RGBA", (W, H), bytes(img.constBits()), "raw", "BGRA", 0, 1)

def load_image(job, src, cb):
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None
    e = ext_of(src.name); err = None
    if has("pillow_heif"):
        try:
            import pillow_heif; pillow_heif.register_heif_opener()
        except Exception: pass
    if e in ("svg", "svgz"):
        try: return svg_to_image(src)
        except Exception as ex: err = ex
    try:
        im = Image.open(src); im.load() if e in ("psd",) else None; return im
    except Exception as ex: err = err or ex
    job.check(); td = Path(tempfile.mkdtemp(prefix="convi_"))
    try:
        ff, _ = find_ffmpeg()
        if ff:
            out = td / "f.png"
            subprocess.run([ff, "-hide_banner", "-nostdin", "-y", "-i", str(src), "-frames:v", "1", str(out)], capture_output=True, creationflags=NOWIN, timeout=120)
            if out.exists(): im = Image.open(out); im.load(); return im
        if S.get("office", True) and find_soffice():
            out = td / "o.png"; lo_convert(job, src, out, "PNG", cb, ("png", "png:draw_png_Export"))
            if out.exists(): im = Image.open(out); im.load(); return im
    except Cancelled: raise
    except Exception as ex: err = ex
    finally: shutil.rmtree(td, ignore_errors=True)
    if e in ("heic", "heif", "hif"): raise RuntimeError("Для HEIC нужна библиотека pillow-heif (ставится при запуске с интернетом)")
    if e in NEED_OFFICE or e in ("eps", "ps", "wmf", "emf"): raise RuntimeError(f"Для .{e} нужен LibreOffice")
    raise RuntimeError(f"Не удалось открыть изображение: {short_err(err)}")

def flat(x):
    from PIL import Image
    if x.mode in ("RGB", "L"): return x
    if x.mode in ("RGBA", "LA") or (x.mode == "P" and "transparency" in x.info):
        x = x.convert("RGBA"); bg = Image.new("RGB", x.size, (255, 255, 255)); bg.paste(x, mask=x.split()[3]); return bg
    return x.convert("RGB")

def conv_image(job, src, dst, fmt, cb):
    from PIL import Image, ImageOps, ImageSequence
    im = load_image(job, src, cb)
    cb(.15); job.check()
    animated = getattr(im, "is_animated", False) and getattr(im, "n_frames", 1) > 1
    if fmt in ("GIF", "WEBP") and animated:
        frames, durs = [], []
        for f in ImageSequence.Iterator(im):
            frames.append(f.convert("RGBA")); durs.append(f.info.get("duration", 100))
        if fmt == "GIF": frames[0].save(dst, "GIF", save_all=True, append_images=frames[1:], duration=durs, loop=0, disposal=2)
        else: frames[0].save(dst, "WEBP", save_all=True, append_images=frames[1:], duration=durs, loop=0, quality=88)
        return
    try: im.seek(0)
    except Exception: pass
    try: im = ImageOps.exif_transpose(im)
    except Exception: pass
    cb(.4)
    if fmt == "JPG": flat(im).save(dst, "JPEG", quality=S.get("quality", 90), optimize=True)
    elif fmt == "PNG":
        if im.mode not in ("1", "L", "LA", "P", "RGB", "RGBA", "I;16"): im = im.convert("RGBA")
        im.save(dst, "PNG", optimize=True)
    elif fmt == "WEBP":
        if im.mode not in ("RGB", "RGBA"): im = im.convert("RGBA" if "A" in im.mode or "transparency" in im.info else "RGB")
        im.save(dst, "WEBP", quality=S.get("quality", 90), method=4)
    elif fmt == "AVIF": im.convert("RGBA" if "A" in im.mode else "RGB").save(dst, "AVIF", quality=max(40, S.get("quality", 90) - 10))
    elif fmt == "BMP": flat(im).save(dst, "BMP")
    elif fmt == "GIF":
        if im.mode not in ("L", "P", "RGB", "RGBA"): im = im.convert("RGB")
        im.save(dst, "GIF")
    elif fmt == "TIFF":
        if im.mode not in ("1", "L", "LA", "P", "RGB", "RGBA", "CMYK"): im = im.convert("RGBA")
        im.save(dst, "TIFF", compression="tiff_lzw")
    elif fmt == "ICO":
        im = im.convert("RGBA"); side = max(im.size)
        cv = Image.new("RGBA", (side, side), (0, 0, 0, 0)); cv.paste(im, ((side - im.width) // 2, (side - im.height) // 2))
        if side > 256: cv = cv.resize((256, 256), Image.LANCZOS); side = 256
        sizes = [(s, s) for s in (16, 24, 32, 48, 64, 128, 256) if s <= side] or [(side, side)]
        cv.save(dst, "ICO", sizes=sizes)
    elif fmt == "PDF": flat(im).save(dst, "PDF", resolution=float(S.get("dpi", 150)))
    else: raise RuntimeError("Неизвестный формат " + fmt)
    cb(1)

def av_args(fmt, enc, kind):
    v264 = "libx264" if "libx264" in enc else "mpeg4"
    vf_even = ["-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2"]
    mp4v = ["-map", "0:v:0?", "-map", "0:a:0?", "-sn", "-c:v", v264] + (["-preset", "medium", "-crf", "22", "-pix_fmt", "yuv420p"] if v264 == "libx264" else ["-q:v", "3"]) + vf_even + ["-c:a", "aac", "-b:a", "192k"]
    if fmt == "WMV": return ["-map", "0:v:0?", "-map", "0:a:0?", "-sn", "-c:v", "wmv2", "-b:v", "3000k", "-c:a", "wmav2", "-b:a", "192k"]
    if fmt == "AAC": return ["-vn", "-sn", "-c:a", "aac", "-b:a", "192k"]
    if fmt == "WMA": return ["-vn", "-sn", "-c:a", "wmav2", "-b:a", "192k"]
    if fmt == "OPUS":
        if "libopus" not in enc: raise RuntimeError("В этой сборке FFmpeg нет кодека Opus")
        return ["-vn", "-sn", "-c:a", "libopus", "-b:a", "160k"]
    if fmt == "MP4": return mp4v + ["-movflags", "+faststart"]
    if fmt in ("MKV", "MOV"): return mp4v
    if fmt == "AVI": return ["-map", "0:v:0?", "-map", "0:a:0?", "-sn", "-c:v", "mpeg4", "-q:v", "3", "-vtag", "XVID"] + vf_even + ["-c:a", "libmp3lame" if "libmp3lame" in enc else "ac3", "-b:a", "192k"]
    if fmt == "WEBM":
        vp = "libvpx-vp9" if "libvpx-vp9" in enc else "libvpx"
        a = ["-c:a", "libopus" if "libopus" in enc else "libvorbis", "-b:a", "128k"]
        return ["-map", "0:v:0?", "-map", "0:a:0?", "-sn", "-c:v", vp, "-crf", "32", "-b:v", "0"] + (["-row-mt", "1", "-deadline", "good", "-cpu-used", "4"] if vp == "libvpx-vp9" else []) + a
    if fmt == "GIF": return ["-an", "-sn", "-vf", "fps=12,scale=w='min(480,iw)':h=-1:flags=lanczos,split[s0][s1];[s0]palettegen[p];[s1][p]paletteuse", "-loop", "0"]
    if fmt == "MP3":
        if "libmp3lame" not in enc: raise RuntimeError("В этой сборке FFmpeg нет кодека MP3 (libmp3lame)")
        return ["-vn", "-sn", "-c:a", "libmp3lame", "-q:a", "2"]
    if fmt == "WAV": return ["-vn", "-sn", "-c:a", "pcm_s16le"]
    if fmt == "FLAC": return ["-vn", "-sn", "-c:a", "flac"]
    if fmt == "M4A": return ["-vn", "-sn", "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart"]
    if fmt == "OGG":
        if "libvorbis" in enc: return ["-vn", "-sn", "-c:a", "libvorbis", "-q:a", "5"]
        if "libopus" in enc: return ["-vn", "-sn", "-c:a", "libopus", "-b:a", "160k"]
        return ["-vn", "-sn", "-c:a", "vorbis", "-strict", "-2", "-b:a", "160k"]
    raise RuntimeError("Неизвестный формат " + fmt)

def ff_probe(ff, src):
    info = subprocess.run([ff, "-hide_banner", "-i", str(src)], capture_output=True, text=True, encoding="utf-8", errors="replace", creationflags=NOWIN).stderr
    m = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", info)
    return info, (int(m.group(1)) * 3600 + int(m.group(2)) * 60 + float(m.group(3)) if m else 0)

def ff_run(job, cmd, dst, dur, cb):
    p = subprocess.Popen(cmd + ["-progress", "pipe:1", "-nostats", str(dst)], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace", creationflags=NOWIN)
    job.proc = p; tail = []
    def drain():
        for ln in p.stderr:
            ln = ln.strip()
            if ln: tail.append(ln); del tail[:-12]
    th = threading.Thread(target=drain, daemon=True); th.start()
    fake = 0.0
    for ln in p.stdout:
        if job.cancel: break
        if ln.startswith("out_time="):
            mm = re.match(r"out_time=(\d+):(\d+):(\d+(?:\.\d+)?)", ln.strip())
            if mm and dur: cb(min(.99, (int(mm.group(1)) * 3600 + int(mm.group(2)) * 60 + float(mm.group(3))) / dur))
            elif not dur: fake = min(.9, fake + .01); cb(fake)
    p.wait(); th.join(timeout=2); job.proc = None
    if job.cancel: raise Cancelled()
    if p.returncode != 0 or not Path(dst).exists():
        err = [l for l in tail if not l.startswith(("ffmpeg version", "built with", "configuration", "lib"))] or tail
        raise RuntimeError("FFmpeg: " + (" | ".join(err[-2:]) if err else f"код {p.returncode}"))
    cb(1)

def conv_av(job, src, dst, fmt, cb):
    ff, _ = find_ffmpeg()
    if not ff: raise RuntimeError("FFmpeg не найден (установи imageio-ffmpeg: pip install imageio-ffmpeg)")
    info, dur = ff_probe(ff, src)
    if fmt in ("MP3", "WAV", "OGG", "FLAC", "M4A", "AAC", "OPUS", "WMA") and "Stream #" in info and "Audio:" not in info: raise RuntimeError("В файле нет звуковой дорожки")
    if fmt in ("MP4", "AVI", "MKV", "MOV", "WEBM", "WMV", "GIF") and "Stream #" in info and "Video:" not in info: raise RuntimeError("В файле нет видео")
    ff_run(job, [ff, "-hide_banner", "-nostdin", "-y", "-i", str(src)] + av_args(fmt, ff_encoders(ff), None), dst, dur, cb)

def make_thumb(path, kind, size=260):
    try:
        from PIL import Image, ImageOps
        if kind == "image":
            if has("pillow_heif"):
                import pillow_heif; pillow_heif.register_heif_opener()
            im = Image.open(path); im.draft("RGB", (size * 2, size * 2)); im = ImageOps.exif_transpose(im)
        elif kind == "video":
            ff, _ = find_ffmpeg()
            if not ff: return None
            td = Path(tempfile.mkdtemp(prefix="convi_t_"))
            try:
                info, dur = ff_probe(ff, path); out = td / "t.png"
                subprocess.run([ff, "-hide_banner", "-nostdin", "-y", "-ss", "1" if dur > 3 else "0", "-i", str(path), "-frames:v", "1", str(out)], capture_output=True, creationflags=NOWIN, timeout=30)
                if not out.exists(): return None
                im = Image.open(out); im.load()
            finally: shutil.rmtree(td, ignore_errors=True)
        else: return None
        b = io.BytesIO(); ImageOps.fit(im.convert("RGB"), (size, size), Image.LANCZOS).save(b, "PNG"); return b.getvalue()
    except Exception: return None

def R(t, b=False, i=False): return [(t, b, i)]
def runs_text(runs): return "".join(r[0] for r in runs)

class _HR(HTMLParser):
    BLOCK = {"p", "div", "li", "pre", "blockquote", "section", "article", "tr", "ul", "ol", "table", "h1", "h2", "h3", "h4", "h5", "h6", "body", "header", "footer"}
    def __init__(s):
        super().__init__(convert_charrefs=True)
        s.blocks = []; s.runs = []; s.b = 0; s.i = 0; s.skip = 0; s.kind = ("p",); s.lists = []; s.pre = 0
        s.table = None; s.row = None; s.cell = None
    def flush(s):
        runs = []
        for t, b, i in s.runs:
            if runs and runs[-1][1:] == (b, i): runs[-1] = (runs[-1][0] + t, b, i)
            else: runs.append((t, b, i))
        if runs:
            runs[0] = (runs[0][0].lstrip(), runs[0][1], runs[0][2]); runs[-1] = (runs[-1][0].rstrip(), runs[-1][1], runs[-1][2])
            runs = [r for r in runs if r[0]]
        s.runs = []
        if not runs: return
        k = s.kind
        if k[0] == "h": s.blocks.append(("h", k[1], runs))
        elif k[0] == "li": s.blocks.append(("li", runs, k[1]))
        else: s.blocks.append(("p", runs))
    def handle_starttag(s, tag, a):
        if tag in ("script", "style", "head", "title", "noscript"): s.skip += 1; return
        if s.skip: return
        if tag == "table": s.flush(); s.table = []; return
        if s.table is not None:
            if tag == "tr": s.row = []
            elif tag in ("td", "th"): s.cell = []
            elif tag == "br" and s.cell is not None: s.cell.append(" ")
            return
        if tag in ("ul", "ol"): s.flush(); s.lists.append(tag == "ol"); return
        if tag in s.BLOCK:
            s.flush()
            if tag[0] == "h" and len(tag) == 2 and tag[1].isdigit(): s.kind = ("h", int(tag[1]))
            elif tag == "li": s.kind = ("li", bool(s.lists and s.lists[-1]))
            else: s.kind = ("p",)
            if tag == "pre": s.pre += 1
        elif tag in ("b", "strong"): s.b += 1
        elif tag in ("i", "em"): s.i += 1
        elif tag == "br": s.runs.append(("\n", 0, 0))
        elif tag == "hr": s.flush()
    def handle_endtag(s, tag):
        if tag in ("script", "style", "head", "title", "noscript"): s.skip = max(0, s.skip - 1); return
        if s.skip: return
        if tag == "table":
            if s.table: s.blocks.append(("table", s.table))
            s.table = None; s.kind = ("p",); return
        if s.table is not None:
            if tag in ("td", "th") and s.cell is not None and s.row is not None: s.row.append(re.sub(r"\s+", " ", "".join(s.cell)).strip()); s.cell = None
            elif tag == "tr" and s.row is not None:
                if s.row: s.table.append(s.row)
                s.row = None
            return
        if tag in ("ul", "ol"):
            s.flush()
            if s.lists: s.lists.pop()
            s.kind = ("p",); return
        if tag in s.BLOCK:
            s.flush(); s.kind = ("li", bool(s.lists and s.lists[-1])) if (tag != "li" and s.lists and tag in ("p", "div")) else ("p",)
            if tag == "pre": s.pre = max(0, s.pre - 1)
        elif tag in ("b", "strong"): s.b = max(0, s.b - 1)
        elif tag in ("i", "em"): s.i = max(0, s.i - 1)
    def handle_data(s, d):
        if s.skip: return
        if s.cell is not None: s.cell.append(d); return
        if s.table is not None: return
        if not s.pre: d = re.sub(r"\s+", " ", d)
        if d: s.runs.append((d, bool(s.b), bool(s.i)))

def read_html_text(txt):
    p = _HR(); p.feed(txt); p.flush(); return p.blocks

def decode_bytes(raw):
    for enc in ("utf-8-sig", "utf-16", "cp1251", "cp1252"):
        try:
            if enc == "utf-16" and not raw.startswith((b"\xff\xfe", b"\xfe\xff")): continue
            return raw.decode(enc)
        except Exception: continue
    return raw.decode("utf-8", "replace")

def read_txt(p):
    t = decode_bytes(Path(p).read_bytes()).replace("\r\n", "\n").replace("\r", "\n")
    return [("p", R(x.strip("\n"))) for x in re.split(r"\n\s*\n", t) if x.strip()]

def read_md(p):
    import markdown
    return read_html_text(markdown.markdown(decode_bytes(Path(p).read_bytes()), extensions=["tables", "fenced_code", "sane_lists"]))

def read_html(p): return read_html_text(decode_bytes(Path(p).read_bytes()))

def read_docx(p):
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = Document(str(p)); blocks = []
    for el in d.element.body.iterchildren():
        tag = el.tag.rsplit("}", 1)[-1]
        if tag == "p":
            para = Paragraph(el, d)
            runs = [(r.text, bool(r.bold), bool(r.italic)) for r in para.runs if r.text]
            if not runs and para.text.strip(): runs = R(para.text)
            if not runs: continue
            st = ""
            try: st = para.style.name or ""
            except Exception: pass
            m = re.match(r"(?:Heading|Заголовок)\s*(\d)", st, re.I)
            if m: blocks.append(("h", int(m.group(1)), runs))
            elif st in ("Title", "Название"): blocks.append(("h", 1, runs))
            elif st.lower().startswith(("list", "список")) or (para._p.pPr is not None and para._p.pPr.numPr is not None):
                blocks.append(("li", runs, "number" in st.lower() or "нумер" in st.lower()))
            else: blocks.append(("p", runs))
        elif tag == "tbl":
            t = Table(el, d)
            rows = [[c.text.strip() for c in r.cells] for r in t.rows]
            if rows: blocks.append(("table", rows))
    return blocks

def read_pdf(p):
    from pypdf import PdfReader
    rd = PdfReader(str(p))
    if rd.is_encrypted:
        try: rd.decrypt("")
        except Exception: raise RuntimeError("PDF защищён паролем")
    blocks = []
    for pg in rd.pages:
        lines = [l.rstrip() for l in (pg.extract_text() or "").splitlines()]
        mx = max([len(l) for l in lines] + [1]); cur = ""
        for l in lines:
            if not l.strip():
                if cur: blocks.append(("p", R(cur))); cur = ""
                continue
            if cur and len(cur) > mx * .72 and not re.search(r"[.!?:;»”)]\s*$", cur):
                cur = cur[:-1] + l.strip() if cur.endswith("-") else cur + " " + l.strip()
            else:
                if cur: blocks.append(("p", R(cur)))
                cur = l.strip()
        if cur: blocks.append(("p", R(cur)))
    if not blocks: raise RuntimeError("В PDF нет текстового слоя (похоже на скан) - для такого нужен OCR")
    return blocks

def read_rtf(p):
    raw = Path(p).read_bytes().decode("latin-1")
    cp = "cp1252"; m = re.search(r"\\ansicpg(\d+)", raw)
    if m: cp = "cp" + m.group(1)
    ignore = {"fonttbl", "colortbl", "stylesheet", "info", "pict", "header", "footer", "footnote", "themedata", "colorschememapping", "latentstyles", "datastore", "generator", "listtable", "listoverridetable", "rsidtbl", "xmlnstbl", "fldinst"}
    tok = re.compile(r"\\([a-zA-Z]+)(-?\d+)? ?|\\'([0-9a-fA-F]{2})|\\(\*)|\\([\\{}~_\-])|([{}])|([^\\{}\r\n]+)|([\r\n]+)")
    out = []; stack = []; skipping = False; uc = 1; skipn = 0
    names = {"emdash": "—", "endash": "–", "bullet": "•", "lquote": "‘", "rquote": "’", "ldblquote": "“", "rdblquote": "”", "tab": "\t"}
    for m in tok.finditer(raw):
        word, num, hx, star, sym, brace, text, nl = m.groups()
        if brace == "{": stack.append((skipping, uc))
        elif brace == "}":
            if stack: skipping, uc = stack.pop()
        elif star: skipping = True
        elif word:
            if word in ignore: skipping = True
            elif skipping: continue
            elif word in ("par", "sect", "page"): out.append("\n")
            elif word == "line": out.append("\n")
            elif word in names: out.append(names[word])
            elif word == "uc" and num: uc = int(num)
            elif word == "u" and num:
                n = int(num); out.append(chr(n + 65536 if n < 0 else n)); skipn = uc
        elif hx and not skipping:
            if skipn: skipn -= 1; continue
            out.append(bytes([int(hx, 16)]).decode(cp, "replace"))
        elif sym and not skipping: out.append({"~": "\u00a0", "_": "-", "-": ""}.get(sym, sym))
        elif text and not skipping:
            if skipn: k = min(skipn, len(text)); text = text[k:]; skipn -= k
            out.append(text)
    t = "".join(out)
    blocks = [("p", R(l.strip())) for l in t.split("\n") if l.strip()]
    if not blocks: raise RuntimeError("Не удалось прочитать RTF")
    return blocks

def read_odt(p):
    import xml.etree.ElementTree as ET
    ns = {"t": "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "o": "urn:oasis:names:tc:opendocument:xmlns:office:1.0", "tb": "urn:oasis:names:tc:opendocument:xmlns:table:1.0"}
    T = lambda n: "{%s}%s" % (ns["t"], n)
    with zipfile.ZipFile(p) as z: root = ET.fromstring(z.read("content.xml"))
    def tx(e):
        o = [e.text or ""]
        for c in e:
            tg = c.tag.rsplit("}", 1)[-1]
            if tg == "s": o.append(" " * int(c.get(T("c"), "1")))
            elif tg == "tab": o.append("\t")
            elif tg == "line-break": o.append("\n")
            else: o.append(tx(c))
            o.append(c.tail or "")
        return "".join(o)
    blocks = []
    def walk(e):
        for c in e:
            tg = c.tag.rsplit("}", 1)[-1]
            if tg == "h":
                t = tx(c).strip()
                if t: blocks.append(("h", int(c.get(T("outline-level"), "1") or 1), R(t)))
            elif tg == "p":
                t = tx(c).strip()
                if t: blocks.append(("p", R(t)))
            elif tg == "list":
                for li in c.iter(T("list-item")):
                    for pp in li.findall(T("p")):
                        t = tx(pp).strip()
                        if t: blocks.append(("li", R(t), False))
            elif tg == "table":
                rows = []
                for tr in c.iter("{%s}table-row" % ns["tb"]):
                    rows.append([" ".join(tx(pp).strip() for pp in cell.iter(T("p"))) for cell in tr.findall("{%s}table-cell" % ns["tb"])])
                if rows: blocks.append(("table", rows))
            elif tg in ("section", "text"): walk(c)
    body = root.find("o:body", ns); walk(body.find("o:text", ns))
    return blocks

CODE_NOTE = "Файл бинарный, а не текстовый - такой формат прочитать нельзя"
def read_code(p):
    raw = Path(p).read_bytes()
    if b"\0" in raw[:8192]: raise RuntimeError(CODE_NOTE)
    t = decode_bytes(raw).replace("\r\n", "\n").replace("\r", "\n").expandtabs(4)
    if not t.strip(): raise RuntimeError("Файл пустой")
    return [("code", t)]

READERS = dict(text=read_txt, log=read_txt, markdown=read_md, txt=read_txt, md=read_md, html=read_html, htm=read_html, docx=read_docx, pdf=read_pdf, rtf=read_rtf, odt=read_odt)

def _he(s): return _html.escape(s).replace("\n", "<br>")
def runs_html(runs):
    o = []
    for t, b, i in runs:
        x = _he(t)
        if b: x = f"<b>{x}</b>"
        if i: x = f"<i>{x}</i>"
        o.append(x)
    return "".join(o)

def w_html(blocks, dst):
    o = ['<!DOCTYPE html>\n<html lang="ru"><head><meta charset="utf-8"><title>%s</title><style>body{font-family:"Segoe UI",Arial,sans-serif;max-width:820px;margin:40px auto;padding:0 20px;line-height:1.6;color:#222}table{border-collapse:collapse}td,th{border:1px solid #bbb;padding:6px 10px}</style></head><body>\n' % _he(dst.stem)]
    lst = None
    for b in blocks:
        if b[0] != "li" and lst: o.append(f"</{lst}>"); lst = None
        if b[0] == "h": n = min(max(b[1], 1), 6); o.append(f"<h{n}>{runs_html(b[2])}</h{n}>")
        elif b[0] == "p": o.append(f"<p>{runs_html(b[1])}</p>")
        elif b[0] == "li":
            tag = "ol" if b[2] else "ul"
            if lst != tag:
                if lst: o.append(f"</{lst}>")
                o.append(f"<{tag}>"); lst = tag
            o.append(f"<li>{runs_html(b[1])}</li>")
        elif b[0] == "table":
            o.append("<table>" + "".join("<tr>" + "".join(f"<td>{_he(c)}</td>" for c in r) + "</tr>" for r in b[1]) + "</table>")
        elif b[0] == "code": o.append('<pre style="background:#f5f5f5;padding:12px;overflow:auto;font:13px Consolas,monospace">' + _html.escape(b[1]) + "</pre>")
    if lst: o.append(f"</{lst}>")
    o.append("</body></html>")
    Path(dst).write_text("\n".join(o), "utf-8")

def runs_md(runs):
    o = []
    for t, b, i in runs:
        if not t.strip(): o.append(t); continue
        lead = t[:len(t) - len(t.lstrip())]; trail = t[len(t.rstrip()):]; core = t.strip()
        if b and i: core = f"***{core}***"
        elif b: core = f"**{core}**"
        elif i: core = f"*{core}*"
        o.append(lead + core + trail)
    return "".join(o).replace("\n", "  \n")

def w_md(blocks, dst):
    o = []; n = 0; prev = None
    for b in blocks:
        if b[0] == "h": o.append("#" * min(max(b[1], 1), 6) + " " + runs_text(b[2])); n = 0
        elif b[0] == "p": o.append(runs_md(b[1])); n = 0
        elif b[0] == "li":
            n = n + 1 if b[2] else 0
            line = (f"{n}. " if b[2] else "- ") + runs_md(b[1])
            if prev == "li" and o: o[-1] += "\n" + line
            else: o.append(line)
        elif b[0] == "table":
            rows = b[1]; w = max(len(r) for r in rows); rows = [r + [""] * (w - len(r)) for r in rows]
            f = lambda r: "| " + " | ".join(c.replace("|", "\\|") for c in r) + " |"
            o.append("\n".join([f(rows[0]), "|" + " --- |" * w] + [f(r) for r in rows[1:]])); n = 0
        elif b[0] == "code": o.append("```\n" + b[1].rstrip("\n") + "\n```"); n = 0
        prev = b[0]
    Path(dst).write_text("\n\n".join(o) + "\n", "utf-8")

def w_txt(blocks, dst):
    o = []; n = 0; prev = None
    for b in blocks:
        if b[0] == "h": o.append(runs_text(b[2]).upper() if b[1] == 1 else runs_text(b[2])); n = 0
        elif b[0] == "p": o.append(runs_text(b[1])); n = 0
        elif b[0] == "li":
            n = n + 1 if b[2] else 0
            line = (f"{n}. " if b[2] else "• ") + runs_text(b[1])
            if prev == "li" and o: o[-1] += "\n" + line
            else: o.append(line)
        elif b[0] == "table": o.append("\n".join("\t".join(r) for r in b[1])); n = 0
        elif b[0] == "code": o.append(b[1].rstrip("\n")); n = 0
        prev = b[0]
    Path(dst).write_text("\n\n".join(o) + "\n", "utf-8")

def w_docx(blocks, dst):
    from docx import Document
    from docx.shared import Pt
    d = Document()
    d.styles["Normal"].font.name = "Calibri"; d.styles["Normal"].font.size = Pt(11)
    def fill(par, runs):
        for t, b, i in runs:
            r = par.add_run(t); r.bold = b or None; r.italic = i or None
    for b in blocks:
        if b[0] == "h": fill(d.add_heading("", level=min(max(b[1], 1), 9)), b[2])
        elif b[0] == "p": fill(d.add_paragraph(), b[1])
        elif b[0] == "li": fill(d.add_paragraph(style="List Number" if b[2] else "List Bullet"), b[1])
        elif b[0] == "code":
            for line in b[1].rstrip("\n").split("\n"):
                pr = d.add_paragraph(); pr.paragraph_format.space_after = Pt(0); r = pr.add_run(line); r.font.name = "Consolas"; r.font.size = Pt(9)
        elif b[0] == "table":
            rows = b[1]; w = max(len(r) for r in rows)
            t = d.add_table(rows=len(rows), cols=w); t.style = "Table Grid"
            for ri, r in enumerate(rows):
                for ci, c in enumerate(r): t.cell(ri, ci).text = c
    d.save(str(dst))

def rtf_esc(s):
    o = []
    for ch in s:
        c = ord(ch)
        if ch in "\\{}": o.append("\\" + ch)
        elif ch == "\n": o.append("\\line ")
        elif ch == "\t": o.append("\\tab ")
        elif c < 32: continue
        elif c < 128: o.append(ch)
        else:
            us = [c] if c <= 0xFFFF else [0xD800 + ((c - 0x10000) >> 10), 0xDC00 + ((c - 0x10000) & 0x3FF)]
            for u in us: o.append("\\u%d?" % (u - 65536 if u > 32767 else u))
    return "".join(o)

def w_rtf(blocks, dst):
    def rr(runs): return "".join(("{" + ("\\b" if b else "") + ("\\i" if i else "") + " " + rtf_esc(t) + "}") if (b or i) else rtf_esc(t) for t, b, i in runs)
    o = ["{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1{\\fonttbl{\\f0\\fswiss Calibri;}{\\f1\\fmodern Courier New;}}\\f0\\fs22\n"]; n = 0
    for b in blocks:
        if b[0] == "h": o.append("\\pard\\sb240\\sa120{\\b\\fs%d %s}\\par\n" % ({1: 40, 2: 32, 3: 28}.get(b[1], 24), rtf_esc(runs_text(b[2])))); n = 0
        elif b[0] == "p": o.append("\\pard\\sa160 " + rr(b[1]) + "\\par\n"); n = 0
        elif b[0] == "li":
            n = n + 1 if b[2] else 0
            o.append("\\pard\\fi-360\\li720\\sa80 " + (f"{n}.\\tab " if b[2] else "\\u8226?\\tab ") + rr(b[1]) + "\\par\n")
        elif b[0] == "table":
            for r in b[1]: o.append("\\pard\\sa60 " + "\\tab ".join(rtf_esc(c) for c in r) + "\\par\n")
            n = 0
        elif b[0] == "code": o.append("\\pard\\f1\\fs18 " + rtf_esc(b[1].rstrip("\n")) + "\\f0\\fs22\\par\n")
    o.append("}")
    Path(dst).write_bytes("".join(o).encode("ascii", "replace"))

_PDF_FONT = {}
def pdf_fonts():
    if _PDF_FONT: return _PDF_FONT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    win = Path(os.environ.get("WINDIR", "C:\\Windows")) / "Fonts"
    sets = [(win, ["segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"]),
            (win, ["arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"]),
            (win, ["calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"]),
            (Path("/usr/share/fonts/truetype/dejavu"), ["DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"]),
            (Path("/usr/share/fonts/truetype/liberation"), ["LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"]),
            (Path("/Library/Fonts"), ["Arial.ttf", "Arial Bold.ttf", "Arial Italic.ttf", "Arial Bold Italic.ttf"]),
            (Path("/System/Library/Fonts/Supplemental"), ["Arial.ttf", "Arial Bold.ttf", "Arial Italic.ttf", "Arial Bold Italic.ttf"])]
    for base, names in sets:
        if all((base / n).exists() for n in names):
            for key, n in zip(("AV", "AV-B", "AV-I", "AV-BI"), names): pdfmetrics.registerFont(TTFont(key, str(base / n)))
            pdfmetrics.registerFontFamily("AV", normal="AV", bold="AV-B", italic="AV-I", boldItalic="AV-BI")
            _PDF_FONT.update(n="AV", b="AV-B"); return _PDF_FONT
    _PDF_FONT.update(n="Helvetica", b="Helvetica-Bold"); return _PDF_FONT

_MONO = []
def pdf_mono():
    if _MONO: return _MONO[0]
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    win = Path(os.environ.get("WINDIR", "C:\\Windows")) / "Fonts"
    for f in (win / "consola.ttf", win / "cour.ttf", win / "lucon.ttf", Path("/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"), Path("/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf"), Path("/Library/Fonts/Courier New.ttf")):
        if f.exists():
            pdfmetrics.registerFont(TTFont("AVM", str(f))); _MONO.append("AVM"); return "AVM"
    _MONO.append("Courier"); return "Courier"

def w_pdf(blocks, dst):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Preformatted
    import textwrap
    from xml.sax.saxutils import escape
    f = pdf_fonts()
    body = ParagraphStyle("b", fontName=f["n"], fontSize=11, leading=15.5, spaceAfter=6)
    hs = {n: ParagraphStyle(f"h{n}", parent=body, fontName=f["b"], fontSize=sz, leading=sz + 5, spaceBefore=12, spaceAfter=6) for n, sz in {1: 21, 2: 17, 3: 14, 4: 12.5}.items()}
    li = ParagraphStyle("li", parent=body, leftIndent=18, bulletIndent=5, spaceAfter=3)
    cell = ParagraphStyle("c", parent=body, fontSize=9.5, leading=12, spaceAfter=0)
    def mk(runs):
        o = []
        for t, b, i in runs:
            x = escape(t).replace("\n", "<br/>")
            if b: x = f"<b>{x}</b>"
            if i: x = f"<i>{x}</i>"
            o.append(x)
        return "".join(o) or " "
    fl = []; n = 0
    for b in blocks:
        if b[0] == "h": fl.append(Paragraph(mk(b[2]), hs.get(b[1], hs[4]))); n = 0
        elif b[0] == "p": fl.append(Paragraph(mk(b[1]), body)); n = 0
        elif b[0] == "li":
            n = n + 1 if b[2] else 0
            fl.append(Paragraph(mk(b[1]), li, bulletText=(f"{n}." if b[2] else "•")))
        elif b[0] == "code":
            ms = ParagraphStyle("m", fontName=pdf_mono(), fontSize=8, leading=10.2)
            lines = []
            for ln in b[1].rstrip("\n").split("\n"): lines += textwrap.wrap(ln, 98, drop_whitespace=False, replace_whitespace=False, break_long_words=True) or [""]
            for i in range(0, len(lines), 90): fl.append(Preformatted("\n".join(lines[i:i + 90]), ms))
            n = 0
        elif b[0] == "table":
            rows = b[1]; w = max(len(r) for r in rows)
            data = [[Paragraph(escape(c) or " ", cell) for c in r + [""] * (w - len(r))] for r in rows]
            t = Table(data, colWidths=[(A4[0] - 40 * mm) / w] * w)
            t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), .5, colors.HexColor("#999999")), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee"))]))
            fl += [t, Spacer(1, 8)]; n = 0
    if not fl: fl.append(Paragraph(" ", body))
    SimpleDocTemplate(str(dst), pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=20 * mm, bottomMargin=20 * mm, title=Path(dst).stem).build(fl)

WRITERS = dict(PDF=w_pdf, DOCX=w_docx, TXT=w_txt, HTML=w_html, MD=w_md, RTF=w_rtf)

OFFICE_IN = {"docx", "odt", "rtf"} | DOC_LO
LO_FILTER = {"PDF": ("pdf", "pdf:writer_pdf_Export"), "DOCX": ("docx", "docx:MS Word 2007 XML"), "TXT": ("txt", "txt:Text (encoded):UTF8"),
             "HTML": ("html", "html:XHTML Writer File:UTF8"), "RTF": ("rtf", "rtf:Rich Text Format")}
WORD_FMT = {"PDF": 17, "DOCX": 16, "RTF": 6}

def lo_convert(job, src, dst, fmt, cb, flt=None):
    so = find_soffice(); ext, flt = flt or LO_FILTER[fmt]
    td = Path(tempfile.mkdtemp(prefix="convi_lo_"))
    try:
        cmd = [so, "-env:UserInstallation=" + (td / "profile").as_uri(), "--headless", "--norestore", "--nolockcheck"]
        if src.suffix.lower() == ".pdf": cmd.append("--infilter=writer_pdf_import")
        cmd += ["--convert-to", flt, "--outdir", str(td), str(src)]
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace", creationflags=NOWIN)
        job.proc = p; t0 = time.time()
        while p.poll() is None:
            if job.cancel: p.terminate(); raise Cancelled()
            if time.time() - t0 > 300: p.kill(); raise RuntimeError("LibreOffice: превышено время ожидания")
            cb(min(.9, (time.time() - t0) / 25)); time.sleep(.3)
        job.proc = None
        out = td / (src.stem + "." + ext)
        if not out.exists(): raise RuntimeError("LibreOffice не создал файл: " + (p.stderr.read() or "")[-200:])
        shutil.move(str(out), str(dst))
    finally:
        shutil.rmtree(td, ignore_errors=True)

def ppt_ok():
    if not IS_WIN: return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application")); return has("win32com")
    except Exception: return False

def ppt_convert(job, src, dst, fmt, cb):
    import pythoncom, win32com.client as wc
    pythoncom.CoInitialize(); pp = None
    try:
        pp = wc.DispatchEx("PowerPoint.Application"); cb(.2)
        pr = pp.Presentations.Open(str(src), ReadOnly=True, WithWindow=False); cb(.5)
        pr.SaveAs(str(dst), {"PDF": 32, "PPTX": 24}[fmt]); pr.Close()
    finally:
        try:
            if pp: pp.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()

def word_convert(job, src, dst, fmt, cb):
    import pythoncom, win32com.client as wc
    pythoncom.CoInitialize(); word = None
    try:
        word = wc.DispatchEx("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
        cb(.2)
        doc = word.Documents.Open(str(src), ReadOnly=True, ConfirmConversions=False, AddToRecentFiles=False)
        cb(.5); doc.SaveAs2(str(dst), FileFormat=WORD_FMT[fmt]); doc.Close(False)
    finally:
        try:
            if word: word.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()

def office_to(job, src, dst, fmt, cb, lo_flt=None, word_fmt=True):
    err = None; ext = ext_of(src.name)
    order = []
    if S.get("office", True):
        if word_fmt and fmt in WORD_FMT and ext in (OFFICE_IN | {"doc"}) and word_ok(): order.append("word")
        if lo_flt is not None or fmt in LO_FILTER:
            if find_soffice(): order.append("libre")
    for e in order:
        try:
            job.check(); (word_convert(job, src, dst, fmt, cb) if e == "word" else lo_convert(job, src, dst, fmt, cb, lo_flt)); return None
        except Cancelled: raise
        except Exception as ex:
            err = ex
            try: Path(dst).unlink()
            except Exception: pass
    return err or False

def conv_doc(job, src, dst, fmt, cb):
    ext = ext_of(src.name)
    pdf_docx = ext == "pdf" and fmt == "DOCX"
    if ext in OFFICE_IN or pdf_docx:
        r = office_to(job, src, dst, fmt, cb)
        if r is None: cb(1); return
        if ext in DOC_LO:
            if not S.get("office", True): raise RuntimeError(f"Для .{ext} включи в настройках «Использовать Word / LibreOffice»")
            if r is False and not (word_ok() or find_soffice()): raise RuntimeError(f"Для .{ext} нужен LibreOffice или Microsoft Word")
            tmpd = Path(tempfile.mkdtemp(prefix="convi_doc_")); mid = tmpd / (src.stem + ".docx")
            try:
                if office_to(job, src, mid, "DOCX", lambda x: cb(x * .5)) is not None: raise RuntimeError(f"Не удалось открыть .{ext}" + (f": {short_err(r)}" if r else ""))
                blocks = read_docx(mid)
            finally: shutil.rmtree(tmpd, ignore_errors=True)
            job.check(); WRITERS[fmt](blocks, Path(dst)); cb(1); return
    cb(.2); blocks = (READERS.get(ext) or read_code)(src)
    job.check(); cb(.6); WRITERS[fmt](blocks, Path(dst)); cb(1)

import csv, datetime
LO_TABLE = {"XLSX": ("xlsx", "xlsx:Calc MS Excel 2007 XML"), "CSV": ("csv", "csv:Text - txt - csv (StarCalc):59,34,76,1"),
            "PDF": ("pdf", "pdf:calc_pdf_Export"), "HTML": ("html", "html:HTML (StarCalc)")}

def _rows(rows):
    out = []
    for r in rows:
        r = ["" if v is None else v for v in r]
        while r and r[-1] == "": r.pop()
        out.append(r)
    while out and not out[-1]: out.pop()
    return out

def _txt(v):
    if isinstance(v, datetime.datetime): return v.date().isoformat() if not (v.hour or v.minute or v.second) else v.isoformat(" ")
    if isinstance(v, (datetime.date, datetime.time)): return v.isoformat()
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15: return str(int(v))
    return "" if v is None else str(v)

def read_ods(p):
    import xml.etree.ElementTree as ET
    T = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"; X = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"; O = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
    with zipfile.ZipFile(p) as z: root = ET.fromstring(z.read("content.xml"))
    sheets = []
    for t in root.iter(T + "table"):
        rows = []
        for tr in t.iter(T + "table-row"):
            row = []
            for c in tr:
                if c.tag not in (T + "table-cell", T + "covered-table-cell"): continue
                v = " ".join("".join(pp.itertext()) for pp in c.findall(X + "p"))
                if c.get(O + "value-type") in ("float", "percentage", "currency") and c.get(O + "value"):
                    try: f = float(c.get(O + "value")); v = int(f) if f == int(f) else f
                    except ValueError: pass
                row += [v] * min(int(c.get(T + "number-columns-repeated", "1")), 50)
            rows += [list(row) for _ in range(min(int(tr.get(T + "number-rows-repeated", "1")), 50))]
        sheets.append((t.get(T + "name", "Лист"), _rows(rows)))
    return sheets, False

def read_table(p):
    e = ext_of(p.name)
    if e in ("csv", "tsv"):
        text = decode_bytes(Path(p).read_bytes()); first = next((l for l in text.splitlines() if l.strip()), "")
        d = "\t" if e == "tsv" else max(",;\t|", key=lambda c: first.count(c))
        return [("Лист1", _rows(csv.reader(io.StringIO(text), delimiter=d)))], True
    if e in ("xlsx", "xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(str(p), data_only=True, read_only=True)
        try: return [(ws.title, _rows(ws.iter_rows(values_only=True))) for ws in wb.worksheets], False
        finally: wb.close()
    if e == "xls":
        import xlrd
        wb = xlrd.open_workbook(str(p)); out = []
        for sh in wb.sheets():
            rows = []
            for r in range(sh.nrows):
                row = []
                for c in range(sh.ncols):
                    cl = sh.cell(r, c); v = cl.value
                    if cl.ctype == xlrd.XL_CELL_DATE:
                        try: v = xlrd.xldate.xldate_as_datetime(v, wb.datemode)
                        except Exception: pass
                    elif cl.ctype == xlrd.XL_CELL_NUMBER and v == int(v): v = int(v)
                    elif cl.ctype == xlrd.XL_CELL_BOOLEAN: v = bool(cl.value)
                    elif cl.ctype == xlrd.XL_CELL_ERROR: v = ""
                    row.append(v)
                rows.append(row)
            out.append((sh.name, _rows(rows)))
        return out, False
    if e == "ods": return read_ods(p)
    raise ImportError("нет читателя для ." + e)

def _num(v):
    if isinstance(v, str):
        s = v.strip()
        if re.fullmatch(r"-?(0|[1-9]\d{0,14})", s): return int(s)
        if re.fullmatch(r"-?(0|[1-9]\d*)[.,]\d+", s): return float(s.replace(",", "."))
    return v

def t_xlsx(sheets, dst, guess):
    import openpyxl
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook(); wb.remove(wb.active); used = set()
    for name, rows in sheets or [("Лист1", [])]:
        title = re.sub(r"[\[\]:*?/\\]", "_", str(name))[:31] or "Лист"; k = 1
        while title.lower() in used: k += 1; title = (title[:28] + f"_{k}")
        used.add(title.lower()); ws = wb.create_sheet(title)
        for r in rows: ws.append([_num(v) if guess else v for v in r])
        for i in range(1, max([len(r) for r in rows] + [1]) + 1):
            ws.column_dimensions[get_column_letter(i)].width = min(60, max([len(_txt(r[i - 1])) for r in rows if len(r) >= i] + [4]) + 2)
    wb.save(str(dst))

def t_csv(sheets, dst, guess):
    with open(dst, "w", encoding="utf-8-sig", newline="") as f: csv.writer(f, delimiter=";").writerows([[_txt(v) for v in r] for r in sheets[0][1]])

def t_html(sheets, dst, guess):
    o = ['<!DOCTYPE html>\n<html lang="ru"><head><meta charset="utf-8"><style>body{font-family:"Segoe UI",Arial,sans-serif;margin:30px}table{border-collapse:collapse;margin-bottom:28px}td,th{border:1px solid #bbb;padding:5px 10px}th{background:#eee}</style></head><body>']
    for name, rows in sheets:
        if len(sheets) > 1: o.append(f"<h2>{_he(str(name))}</h2>")
        o.append("<table>" + "".join("<tr>" + "".join(f"<{'th' if i == 0 else 'td'}>{_he(_txt(c))}</{'th' if i == 0 else 'td'}>" for c in r) + "</tr>" for i, r in enumerate(rows)) + "</table>")
    Path(dst).write_text("\n".join(o) + "</body></html>", "utf-8")

def t_json(sheets, dst, guess):
    rows = sheets[0][1]; head = rows[0] if rows else []
    if head and all(isinstance(h, str) and h for h in head) and len(set(head)) == len(head):
        data = [{h: (r[i] if i < len(r) else "") for i, h in enumerate(head)} for r in rows[1:]]
    else: data = rows
    Path(dst).write_text(json.dumps(data, ensure_ascii=False, indent=1, default=_txt), "utf-8")

def t_pdf(sheets, dst, guess):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape
    f = pdf_fonts(); mc = max([len(r) for _, rows in sheets for r in rows] + [1]); ps = landscape(A4) if mc > 6 else A4
    fs = 9 if mc <= 6 else 8 if mc <= 10 else 6.5
    cell = ParagraphStyle("c", fontName=f["n"], fontSize=fs, leading=fs + 2.5); hd = ParagraphStyle("h", fontName=f["b"], fontSize=14, spaceAfter=6)
    avail = ps[0] - 24 * mm; fl = []
    for name, rows in sheets:
        if len(sheets) > 1: fl.append(Paragraph(escape(str(name)), hd))
        if not rows: continue
        w = max(len(r) for r in rows); wt = [min(40, max([len(_txt(r[i])) for r in rows if len(r) > i] + [3])) + 2 for i in range(w)]
        data = [[Paragraph(escape(_txt(c)) or " ", cell) for c in r + [""] * (w - len(r))] for r in rows]
        t = Table(data, colWidths=[avail * x / sum(wt) for x in wt], repeatRows=1)
        t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), .4, colors.HexColor("#999999")), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee"))]))
        fl += [t, Spacer(1, 10)]
    SimpleDocTemplate(str(dst), pagesize=ps, leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm, bottomMargin=12 * mm).build(fl or [Paragraph(" ", cell)])

TABLE_W = dict(XLSX=t_xlsx, CSV=t_csv, HTML=t_html, JSON=t_json, PDF=t_pdf)

def conv_table(job, src, dst, fmt, cb):
    err = None
    try:
        cb(.1); sheets, guess = read_table(src); job.check(); cb(.5)
        TABLE_W[fmt](sheets, dst, guess); cb(1); return
    except Cancelled: raise
    except Exception as e: err = e
    if fmt in LO_TABLE and S.get("office", True) and find_soffice():
        lo_convert(job, src, dst, fmt, cb, LO_TABLE[fmt]); cb(1); return
    ext = ext_of(src.name)
    if isinstance(err, ImportError) or ext in NEED_OFFICE:
        raise RuntimeError(f"Для .{ext} нужна библиотека " + {"xls": "xlrd", "xlsx": "openpyxl", "xlsm": "openpyxl"}.get(ext, "LibreOffice") + " или LibreOffice")
    raise RuntimeError(short_err(err))

LO_PRES = {"PDF": ("pdf", "pdf:impress_pdf_Export"), "PPTX": ("pptx", "pptx:Impress MS PowerPoint 2007 XML")}

def read_pptx(p):
    from pptx import Presentation
    blocks = []
    for i, sl in enumerate(Presentation(str(p)).slides, 1):
        ts = sl.shapes.title; title = ts.text_frame.text.strip() if ts is not None and ts.has_text_frame else ""
        blocks.append(("h", 1, R(f"{i}. {title}" if title else f"Слайд {i}")))
        for sh in sl.shapes:
            if ts is not None and sh.shape_id == ts.shape_id: continue
            if sh.has_text_frame:
                for pa in sh.text_frame.paragraphs:
                    t = (pa.text or "").strip()
                    if t: blocks.append(("li", R(t), False) if pa.level > 0 else ("p", R(t)))
            elif getattr(sh, "has_table", False) and sh.has_table: blocks.append(("table", [[c.text.strip() for c in r.cells] for r in sh.table.rows]))
    if not blocks: raise RuntimeError("В презентации нет текста")
    return blocks

def conv_pres(job, src, dst, fmt, cb):
    ext = ext_of(src.name); use = S.get("office", True); err = None
    if use and fmt in ("PDF", "PPTX"):
        for e in (["ppt"] if ppt_ok() else []) + (["libre"] if find_soffice() else []):
            try:
                job.check(); (ppt_convert(job, src, dst, fmt, cb) if e == "ppt" else lo_convert(job, src, dst, fmt, cb, LO_PRES[fmt])); cb(1); return
            except Cancelled: raise
            except Exception as ex:
                err = ex
                try: Path(dst).unlink()
                except Exception: pass
    if fmt in WRITERS and ext == "pptx":
        cb(.3); blocks = read_pptx(src); job.check(); WRITERS[fmt](blocks, Path(dst)); cb(1); return
    if fmt in ("TXT", "DOCX") and use and ext != "pptx" and (ppt_ok() or find_soffice()):
        tmpd = Path(tempfile.mkdtemp(prefix="convi_ppt_")); mid = tmpd / (src.stem + ".pptx")
        try:
            for e in (["ppt"] if ppt_ok() else []) + (["libre"] if find_soffice() else []):
                try: (ppt_convert(job, src, mid, "PPTX", cb) if e == "ppt" else lo_convert(job, src, mid, "PPTX", cb, LO_PRES["PPTX"])); break
                except Cancelled: raise
                except Exception as ex: err = ex
            if not mid.exists(): raise RuntimeError("Не удалось открыть презентацию: " + short_err(err))
            blocks = read_pptx(mid)
        finally: shutil.rmtree(tmpd, ignore_errors=True)
        WRITERS[fmt](blocks, Path(dst)); cb(1); return
    raise RuntimeError(f"Для .{ext} в {fmt} нужен Microsoft PowerPoint или LibreOffice" + (f" ({short_err(err)})" if err else ""))

def conv_pdf_images(job, src, dst, fmt, cb):
    try: import fitz
    except ImportError: raise RuntimeError("Для PDF в JPG/PNG нужна библиотека PyMuPDF (ставится при запуске с интернетом)")
    doc = fitz.open(str(src))
    if doc.needs_pass: raise RuntimeError("PDF защищён паролем")
    n = doc.page_count; ext = fmt.lower()
    if n == 0: raise RuntimeError("В PDF нет страниц")
    k = S.get("dpi", 150) / 72.0; mat = fitz.Matrix(k, k)
    def save(i, path):
        job.check(); pix = doc[i].get_pixmap(matrix=mat, alpha=False)
        pix.save(str(path), jpg_quality=92) if ext == "jpg" else pix.save(str(path))
    if n == 1: save(0, dst); cb(1); return None
    folder = unique(Path(dst).parent / (src.stem + " (страницы)")); folder.mkdir()
    for i in range(n): save(i, folder / f"{src.stem}_{i + 1:03d}.{ext}"); cb((i + 1) / n)
    return folder

IMG_OPS = {"compress": "_compressed", "resize": "_resized", "rotate": "_rotated", "bw": "_bw", "strip": "_clean", "watermark": "_wm", "crop": "_crop"}
AV_OPS = {"vid_compress": "_compressed", "vid_mute": "_mute", "trim": "_trim"}
PDF_OPS = {"pdf_rotate": "_rotated", "pdf_compress": "_compressed"}
GROUP_OPS = {"images_pdf", "pdf_merge", "pdf_split"}
KEEP_NAME_OPS = set(IMG_OPS) | set(PDF_OPS) | set(AV_OPS)
WRITE_EXT = {"jpg": "jpg", "jpeg": "jpeg", "jfif": "jpg", "png": "png", "webp": "webp", "avif": "avif", "gif": "gif", "bmp": "bmp", "tif": "tif", "tiff": "tiff"}
LOSSY = ("jpg", "jpeg", "webp", "avif")

def out_dir(src):
    if S.get("mode") == "folder": return Path(S.get("outdir_path") or default_outdir())
    return Path(src).parent

def save_img(im, dst, ext, q=None):
    q = q or S.get("quality", 90); ext = ext.lower()
    if ext in ("jpg", "jpeg"): flat(im).save(dst, "JPEG", quality=q, optimize=True)
    elif ext == "webp": im.convert("RGBA" if "A" in im.mode else "RGB").save(dst, "WEBP", quality=q, method=4)
    elif ext == "avif": im.convert("RGBA" if "A" in im.mode else "RGB").save(dst, "AVIF", quality=q)
    elif ext == "png": im.save(dst, "PNG", optimize=True)
    elif ext == "gif": im.convert("RGB").save(dst, "GIF", optimize=True)
    elif ext == "bmp": flat(im).save(dst, "BMP")
    else: im.save(dst, "TIFF", compression="tiff_lzw")

def encode(im, ext, q=None):
    b = io.BytesIO(); save_img(im, b, ext, q); return b.getvalue()

def fit_target(im, ext, target, cb=None):
    from PIL import Image
    cur = im; scale = 1.0; pal = False
    while True:
        if ext in LOSSY:
            lo, hi, best = 5, 95, None
            while lo <= hi:
                mid = (lo + hi) // 2; b = encode(cur, ext, mid)
                if len(b) <= target: best = b; lo = mid + 1
                else: hi = mid - 1
            if best: return best
        else:
            b = encode(cur, ext)
            if len(b) <= target: return b
            if ext == "png" and not pal:
                pal = True; cur = cur.convert("RGBA" if "A" in cur.mode else "RGB").quantize(256, method=2 if "A" in cur.mode else 0)
                b = encode(cur, ext)
                if len(b) <= target: return b
        scale *= 0.85
        if scale < 0.08: raise RuntimeError("Не удалось уместить в заданный размер")
        cur = im.resize((max(1, int(im.width * scale)), max(1, int(im.height * scale))), Image.LANCZOS)

def img_op(job, src, dst, op, p, cb):
    from PIL import Image, ImageOps
    im = load_image(job, src, cb); ext = dst.suffix.lstrip(".").lower()
    try: im = ImageOps.exif_transpose(im)
    except Exception: pass
    cb(.3); job.check()
    if op == "compress":
        if p.get("target_kb"): dst.write_bytes(fit_target(im, ext, int(p["target_kb"]) * 1024))
        else:
            q = int(p.get("quality", 75))
            if ext == "png" and q < 90: im = im.convert("RGBA" if "A" in im.mode else "RGB").quantize(256, method=2 if "A" in im.mode else 0)
            dst.write_bytes(encode(im, ext, q))
    elif op == "resize":
        w, h = im.size
        if p.get("mode") == "percent": k = float(p.get("value", 50)) / 100
        else:
            bw, bh = int(p.get("w") or 0), int(p.get("h") or 0)
            if not bw and not bh: raise RuntimeError("Укажи ширину или высоту")
            k = min(bw / w if bw else 1e9, bh / h if bh else 1e9)
        im = im.resize((max(1, round(w * k)), max(1, round(h * k))), Image.LANCZOS); save_img(im, dst, ext)
    elif op == "rotate":
        m = {"cw90": Image.Transpose.ROTATE_270, "ccw90": Image.Transpose.ROTATE_90, "180": Image.Transpose.ROTATE_180,
             "fliph": Image.Transpose.FLIP_LEFT_RIGHT, "flipv": Image.Transpose.FLIP_TOP_BOTTOM}[p.get("angle", "cw90")]
        save_img(im.transpose(m), dst, ext)
    elif op == "bw":
        g = ImageOps.grayscale(im); save_img(g if ext != "gif" else g.convert("RGB"), dst, ext)
    elif op == "strip": save_img(im, dst, ext, 95)
    elif op == "crop":
        a, b = (int(x) for x in str(p.get("ratio", "1:1")).split(":")); w, h = im.size
        nw, nh = (w, round(w * b / a)) if w * b <= h * a else (round(h * a / b), h)
        l, t = (w - nw) // 2, (h - nh) // 2; save_img(im.crop((l, t, l + nw, t + nh)), dst, ext)
    elif op == "watermark": save_img(watermark(im, p), dst, ext)
    cb(1)

def find_font():
    win = Path(os.environ.get("WINDIR", "C:\\Windows")) / "Fonts"
    for f in (win / "segoeui.ttf", win / "arial.ttf", Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"), Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"), Path("/Library/Fonts/Arial.ttf")):
        if f.exists(): return str(f)

def watermark(im, p):
    from PIL import Image, ImageDraw, ImageFont
    text = str(p.get("text") or "").strip()
    if not text: raise RuntimeError("Введи текст подписи")
    base = im.convert("RGBA"); W, H = base.size; size = max(12, int(min(W, H) * float(p.get("size", 5)) / 100))
    fp = find_font(); font = ImageFont.truetype(fp, size) if fp else ImageFont.load_default(size)
    al = int(255 * float(p.get("opacity", 60)) / 100); layer = Image.new("RGBA", (W, H), (0, 0, 0, 0)); d = ImageDraw.Draw(layer)
    bb = d.textbbox((0, 0), text, font=font); tw, th = bb[2] - bb[0], bb[3] - bb[1]; m = max(8, size // 2)
    def put(dr, x, y): dr.text((x + 1, y + 1), text, font=font, fill=(0, 0, 0, int(al * .6))); dr.text((x, y), text, font=font, fill=(255, 255, 255, al))
    pos = p.get("pos", "br")
    if pos == "tile":
        tile = Image.new("RGBA", (tw + size * 2, th + size * 2), (0, 0, 0, 0)); put(ImageDraw.Draw(tile), size, size)
        tile = tile.rotate(30, expand=True, resample=Image.BICUBIC)
        for y in range(-tile.height, H + tile.height, int(tile.height * 1.1)):
            for x in range(-tile.width, W + tile.width, int(tile.width * 1.1)): layer.alpha_composite(tile, (x, y)) if 0 <= x < W - 1 and 0 <= y < H - 1 else None
    else:
        x = {"l": m, "r": W - tw - m, "c": (W - tw) // 2}[{"br": "r", "bl": "l", "tr": "r", "tl": "l", "center": "c"}[pos]]
        y = {"b": H - th - m * 2, "t": m, "c": (H - th) // 2}[{"br": "b", "bl": "b", "tr": "t", "tl": "t", "center": "c"}[pos]]
        put(d, x - bb[0], y - bb[1])
    return Image.alpha_composite(base, layer)

def tparse(s):
    s = str(s or "").strip().replace(",", ".")
    if not s: return None
    parts = s.split(":")
    try: v = [float(x) for x in parts]
    except ValueError: raise RuntimeError(f"Неверное время: {s}")
    if len(v) > 3: raise RuntimeError(f"Неверное время: {s}")
    return sum(x * 60 ** i for i, x in enumerate(reversed(v)))

def av_op(job, src, dst, op, p, cb):
    ff, _ = find_ffmpeg()
    if not ff: raise RuntimeError("FFmpeg не найден")
    info, dur = ff_probe(ff, src); base = [ff, "-hide_banner", "-nostdin", "-y"]; hasv = "Video:" in info; hasa = "Audio:" in info
    if op == "vid_compress":
        if not hasv: raise RuntimeError("В файле нет видео")
        if p.get("target_mb"):
            if not dur: raise RuntimeError("Не удалось узнать длительность видео")
            kb = int(float(p["target_mb"]) * 8192 * 0.93 / dur) - 128
            if kb < 120: raise RuntimeError("Слишком мало для такого видео - выбери размер побольше")
            cap = 480 if kb < 700 else 720 if kb < 2500 else 1080 if kb < 6000 else 0
            v = ["-b:v", f"{kb}k", "-maxrate", f"{int(kb * 1.4)}k", "-bufsize", f"{kb * 2}k"]
        else:
            crf, cap = {"hq": (23, 0), "mid": (28, 720), "low": (32, 480)}.get(p.get("preset", "mid"), (28, 720)); v = ["-crf", str(crf)]
        vf = f"scale=-2:'min({cap},ih)'" if cap else "scale=trunc(iw/2)*2:trunc(ih/2)*2"
        enc = ff_encoders(ff)
        cmd = base + ["-i", str(src), "-map", "0:v:0", "-map", "0:a:0?", "-sn", "-c:v", "libx264" if "libx264" in enc else "mpeg4", "-preset", "medium"] + v + ["-vf", vf, "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "128k", "-movflags", "+faststart"]
        ff_run(job, cmd, dst, dur, cb)
    elif op == "vid_mute":
        if not hasv: raise RuntimeError("В файле нет видео")
        if not hasa: raise RuntimeError("В видео и так нет звука")
        ff_run(job, base + ["-i", str(src), "-map", "0:v", "-map", "0:s?", "-c", "copy", "-an"], dst, dur, cb)
    elif op == "trim":
        s = tparse(p.get("start")) or 0.0; e = tparse(p.get("end"))
        if e is None and not dur: raise RuntimeError("Укажи конец: длительность файла неизвестна")
        if dur and s >= dur: raise RuntimeError(f"Начало за пределами файла (длительность {int(dur // 60)}:{int(dur % 60):02d})")
        e = min(e, dur) if (e is not None and dur) else (e if e is not None else dur)
        if s < 0 or e <= s: raise RuntimeError("Конец должен быть позже начала")
        ext = dst.suffix.lower(); tail = ["-c", "copy", "-avoid_negative_ts", "make_zero"] if p.get("mode") != "exact" else (["-c:v", "libx264", "-preset", "fast", "-crf", "20", "-c:a", "aac", "-b:a", "192k"] if hasv else [])
        ff_run(job, base + ["-ss", f"{s:.3f}", "-i", str(src), "-t", f"{e - s:.3f}"] + tail, dst, e - s, cb)
    else: raise RuntimeError("Неизвестный инструмент")

def parse_pages(text, n):
    t = str(text or "").strip().lower()
    if t in ("", "все", "all", "*"): return list(range(n))
    out = []
    for part in re.split(r"[,;\s]+", t):
        if not part: continue
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", part)
        if not m: raise RuntimeError(f"Неверный диапазон страниц: {part}")
        a = int(m[1]); b = int(m[2] or a)
        if a < 1 or b < a or b > n: raise RuntimeError(f"Страницы вне диапазона 1-{n}: {part}")
        out += range(a - 1, b)
    return out

def pdf_open(p):
    from pypdf import PdfReader
    r = PdfReader(str(p))
    if r.is_encrypted:
        try:
            if not r.decrypt(""): raise RuntimeError("PDF защищён паролем")
        except Exception: raise RuntimeError("PDF защищён паролем")
    return r

def pdf_compress(src, dst, level, cb):
    from pypdf import PdfWriter
    w = PdfWriter(clone_from=pdf_open(src)); n = len(w.pages)
    for i, pg in enumerate(w.pages):
        try: pg.compress_content_streams()
        except Exception: pass
        if level >= 2:
            for im in list(pg.images):
                try:
                    pil = im.image
                    if pil.mode not in ("RGB", "L"): pil = pil.convert("RGB")
                    if level >= 3 and max(pil.size) > 1400:
                        k = 1400 / max(pil.size); pil = pil.resize((max(1, int(pil.width * k)), max(1, int(pil.height * k))))
                    im.replace(pil, quality=70 if level == 2 else 45)
                except Exception: pass
        cb((i + 1) / n * .9)
    try: w.compress_identical_objects(remove_identicals=True, remove_orphans=True)
    except Exception: pass
    with open(dst, "wb") as f: w.write(f)

def run_op(job, it, cb):
    from pypdf import PdfWriter
    op = it["op"]; p = it.get("params") or {}; paths = [Path(x) for x in it["paths"]]; src = paths[0]; od = out_dir(src); od.mkdir(parents=True, exist_ok=True)
    size_in = sum(x.stat().st_size for x in paths)
    if op in IMG_OPS:
        dst = unique(od / f"{src.stem}{IMG_OPS[op]}.{WRITE_EXT.get(ext_of(src.name), 'jpg' if ext_of(src.name) in ('heic', 'heif', 'hif') else 'png')}")
        tmp = dst.with_name(f".convi-{uuid.uuid4().hex[:8]}{dst.suffix}"); img_op(job, src, tmp, op, p, cb)
    elif op in AV_OPS:
        e = ext_of(src.name); kind = file_type(src)
        ne = "mp4" if op == "vid_compress" or (op == "trim" and p.get("mode") == "exact" and kind == "video") else e
        dst = unique(od / f"{src.stem}{AV_OPS[op]}.{ne}"); tmp = dst.with_name(f".convi-{uuid.uuid4().hex[:8]}{dst.suffix}"); av_op(job, src, tmp, op, p, cb)
    elif op in ("images_pdf", "pdf_merge"):
        dst = unique(od / f"{src.stem} (+{len(paths) - 1}).pdf"); tmp = dst.with_name(f".convi-{uuid.uuid4().hex[:8]}.pdf")
        if op == "images_pdf":
            ims = []
            for i, x in enumerate(paths): job.check(); ims.append(flat(load_image(job, x, cb))); cb((i + 1) / len(paths) * .9)
            ims[0].save(tmp, "PDF", save_all=True, append_images=ims[1:], resolution=float(S.get("dpi", 150)))
        else:
            w = PdfWriter()
            for i, x in enumerate(paths): job.check(); w.append(pdf_open(x)); cb((i + 1) / len(paths) * .9)
            with open(tmp, "wb") as f: w.write(f)
    elif op == "pdf_rotate":
        r = pdf_open(src); w = PdfWriter(clone_from=r); ang = int(p.get("angle", 90))
        for i in parse_pages(p.get("pages"), len(w.pages)): w.pages[i].rotate(ang)
        dst = unique(od / f"{src.stem}_rotated.pdf"); tmp = dst.with_name(f".convi-{uuid.uuid4().hex[:8]}.pdf")
        with open(tmp, "wb") as f: w.write(f)
    elif op == "pdf_compress":
        dst = unique(od / f"{src.stem}_compressed.pdf"); tmp = dst.with_name(f".convi-{uuid.uuid4().hex[:8]}.pdf")
        pdf_compress(src, tmp, int(p.get("level", 2)), cb)
        if Path(tmp).stat().st_size >= size_in * 0.98: Path(tmp).unlink(); raise RuntimeError("Почти не сжимается - попробуй уровень сильнее")
    elif op == "pdf_split":
        r = pdf_open(src); n = len(r.pages)
        parts = [str(i + 1) for i in range(n)] if p.get("mode") == "each" else [x for x in re.split(r"[,;\s]+", str(p.get("ranges", "")).strip()) if x]
        if not parts: raise RuntimeError("Укажи страницы, например 1-3,5")
        folder = unique(od / f"{src.stem} (части)"); folder.mkdir(parents=True, exist_ok=True)
        for k, part in enumerate(parts):
            job.check(); w = PdfWriter()
            for i in parse_pages(part, n): w.add_page(r.pages[i])
            with open(folder / (f"{src.stem}_{part}.pdf" if p.get("mode") != "each" else f"{src.stem}_стр_{int(part):03d}.pdf"), "wb") as f: w.write(f)
            cb((k + 1) / len(parts))
        return dict(dst=str(folder), size_in=size_in, size_out=sum(x.stat().st_size for x in folder.iterdir()))
    else: raise RuntimeError("Неизвестный инструмент")
    os.replace(tmp, dst)
    return dict(dst=str(dst), size_in=size_in, size_out=Path(dst).stat().st_size)

def convert_file(job, src, fmt, dst, cb):
    t = file_type(src)
    if t == "image": conv_image(job, src, dst, fmt, cb)
    elif t in ("video", "audio"): conv_av(job, src, dst, fmt, cb)
    elif t == "doc": return conv_pdf_images(job, src, dst, fmt, cb) if fmt in ("JPG", "PNG") else conv_doc(job, src, dst, fmt, cb)
    elif t == "table": conv_table(job, src, dst, fmt, cb)
    elif t == "pres": conv_pres(job, src, dst, fmt, cb)
    else: raise RuntimeError("Формат не поддерживается")

HIST_LOCK = threading.Lock()
OP_TITLE = {"compress": "Сжатие", "resize": "Размер", "rotate": "Поворот", "bw": "Ч/Б", "strip": "Без EXIF", "watermark": "Подпись", "crop": "Обрезка",
            "images_pdf": "Фото в PDF", "pdf_merge": "Склейка PDF", "pdf_split": "Разбивка PDF", "pdf_rotate": "Поворот PDF", "pdf_compress": "Сжатие PDF",
            "vid_compress": "Сжатие видео", "vid_mute": "Без звука", "trim": "Обрезка"}

def history_load():
    try: return json.loads((DATA / "history.json").read_text("utf-8"))
    except Exception: return []

def history_add(res, it):
    if not res.get("ok") or res.get("skipped"): return
    try:
        with HIST_LOCK:
            h = history_load(); what = it.get("to") or OP_TITLE.get(it.get("op"), it.get("op", ""))
            h.insert(0, dict(t=int(time.time()), src=Path(res["src"]).name, dst=res["dst"], what=what, si=res.get("size_in", 0), so=res.get("size_out", 0)))
            (DATA / "history.json").write_text(json.dumps(h[:60], ensure_ascii=False), "utf-8")
    except Exception: pass

def history_clear():
    try: (DATA / "history.json").write_text("[]", "utf-8")
    except Exception: pass

def size_of(p):
    p = Path(p)
    return sum(x.stat().st_size for x in p.rglob("*") if x.is_file()) if p.is_dir() else (p.stat().st_size if p.exists() else 0)

def run_job(job):
    total = len(job.items)
    for i, it in enumerate(job.items):
        if job.cancel: break
        job.idx = i; src = Path(it["path"]); job.cur = src.name if "op" not in it or len(it["paths"]) == 1 else f"{src.name} (+{len(it['paths']) - 1})"
        res = dict(src=str(src), ok=False); tmp = None
        cb = lambda p, i=i: setattr(job, "pct", (i + max(0.0, min(1.0, p))) / total)
        try:
            if "op" in it:
                res.update(run_op(job, it, cb)); res["ok"] = True
                if S.get("mode") == "replace" and it["op"] in KEEP_NAME_OPS: replace_original(src, res)
            else:
                fmt = it["to"].upper()
                if not src.is_file(): raise RuntimeError("Файл не найден")
                if same_fmt(ext_of(src.name), fmt): res.update(ok=True, skipped=True, dst=str(src))
                else:
                    folder = out_dir(src); folder.mkdir(parents=True, exist_ok=True)
                    dst = unique(folder / (src.stem + "." + fmt.lower())); tmp = dst.with_name(f".convi-{uuid.uuid4().hex[:8]}{dst.suffix}")
                    ret = convert_file(job, src, fmt, tmp, cb)
                    if ret is not None: dst = Path(ret)
                    else: os.replace(tmp, dst)
                    tmp = None; res.update(ok=True, dst=str(dst), size_in=src.stat().st_size, size_out=size_of(dst))
                    if S.get("mode") == "replace" and dst.exists() and (dst.is_dir() or dst.stat().st_size > 0): replace_original(src, res)
        except Cancelled: res["error"] = "Отменено"
        except Exception as e:
            res["error"] = short_err(e)
            if os.environ.get("CONVI_DEBUG"): traceback.print_exc()
        finally:
            if tmp:
                try: Path(tmp).unlink()
                except Exception: pass
        history_add(res, it); job.results.append(res); job.pct = (i + 1) / total
    job.state = "cancel" if job.cancel else "done"

def replace_original(src, res):
    try:
        if S.get("trash", True):
            from send2trash import send2trash; send2trash(str(src))
        else: src.unlink()
        dst = Path(res["dst"])
        if dst.suffix == src.suffix and dst.is_file() and not src.exists(): os.replace(dst, src); res["dst"] = str(src)
    except Exception as e: res["warn"] = "оригинал не удалён: " + short_err(e)

def start_job(items):
    clean = []
    for it in items:
        if it.get("op"):
            op = it["op"]; ps = [Path(str(x)) for x in (it.get("paths") or [it.get("path")])]
            if not ps or not all(x.is_file() for x in ps): return None, "Файл не найден"
            if op in IMG_OPS or op == "images_pdf":
                if not all(file_type(x) == "image" for x in ps): return None, "Этот инструмент - только для картинок"
            elif op in ("vid_compress", "vid_mute"):
                if not all(file_type(x) == "video" for x in ps): return None, "Этот инструмент - только для видео"
            elif op == "trim":
                if not all(file_type(x) in ("video", "audio") for x in ps): return None, "Обрезать можно видео и аудио"
            elif op in PDF_OPS or op in ("pdf_merge", "pdf_split"):
                if not all(ext_of(x.name) == "pdf" for x in ps): return None, "Этот инструмент - только для PDF"
            else: return None, "Неизвестный инструмент"
            if op in ("images_pdf", "pdf_merge") and len(ps) < 2: return None, "Нужно минимум два файла"
            clean.append(dict(op=op, path=str(ps[0]), paths=[str(x) for x in ps], params=it.get("params") or {})); continue
        p = Path(str(it.get("path", ""))); t = file_type(p); fmt = str(it.get("to", "")).upper()
        if not p.is_file(): return None, f"Файл не найден: {p.name}"
        if not t or fmt not in outputs_for(p.name, t): return None, f"{p.name}: формат {fmt} не поддерживается"
        clean.append(dict(path=str(p), to=fmt))
    if not clean: return None, "Нет файлов"
    job = Job(clean); JOBS[job.id] = job
    threading.Thread(target=run_job, args=(job,), daemon=True).start()
    return job, None

WATCH_KEYS = {"image": "watch_img", "video": "watch_vid", "audio": "watch_aud", "doc": "watch_doc"}
SKIP_EXT = (".tmp", ".part", ".crdownload", ".download", ".partial", ".opdownload")

class Watcher:
    def __init__(self): self.seen = {}; self.pend = {}; self.armed = None
    def tick(self):
        if not S.get("watch_on") or not S.get("watch_dir"): self.armed = None; return []
        d = Path(S["watch_dir"])
        if not d.is_dir(): return []
        cur = []
        for f in d.iterdir():
            try:
                if f.is_file() and not f.name.startswith(".") and f.suffix.lower() not in SKIP_EXT: st = f.stat(); cur.append((f, (st.st_mtime, st.st_size)))
            except OSError: pass
        if self.armed != str(d): self.seen = {str(f): sig for f, sig in cur}; self.armed = str(d); self.pend = {}; return []
        jobs = []
        for f, sig in cur:
            k = str(f)
            if self.seen.get(k) == sig: continue
            if self.pend.get(k) != sig: self.pend[k] = sig; continue
            self.pend.pop(k, None); self.seen[k] = sig
            t = EXT_TYPE.get(ext_of(f.name)); fmt = S.get(WATCH_KEYS.get(t, ""), "") if t else ""
            if not fmt or same_fmt(ext_of(f.name), fmt) or fmt not in outputs_for(f.name, t): continue
            job, err = start_job([dict(path=str(f), to=fmt)])
            if job: jobs.append(job)
        return jobs

def launcher():
    if FROZEN: return f'"{sys.executable}"'
    exe = Path(sys.executable)
    if IS_WIN and exe.with_name("pythonw.exe").exists(): exe = exe.with_name("pythonw.exe")
    return f'"{exe}" "{Path(os.path.abspath(__file__))}"'

def _del_tree(winreg, root, path):
    try: k = winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS)
    except OSError: return
    while True:
        try: sub = winreg.EnumKey(k, 0)
        except OSError: break
        _del_tree(winreg, root, path + "\\" + sub)
    winreg.CloseKey(k)
    try: winreg.DeleteKey(root, path)
    except OSError: pass

def set_autostart(on):
    if not IS_WIN: return "Доступно только в Windows"
    import winreg
    k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
    try:
        if on: winreg.SetValueEx(k, "Convi", 0, winreg.REG_SZ, launcher() + " --tray")
        else:
            try: winreg.DeleteValue(k, "Convi")
            except OSError: pass
    finally: winreg.CloseKey(k)

CTX_EXT = _s("jpg jpeg jfif png webp bmp gif tif tiff ico heic avif svg mp4 avi mkv mov webm wmv flv mpg mpeg 3gp ts mp3 wav ogg flac m4a aac wma opus docx doc pdf txt md html htm rtf odt xlsx xls csv ods pptx ppt odp")

def set_ctx(on):
    if not IS_WIN: return "Доступно только в Windows"
    import winreg
    base = r"Software\Classes\SystemFileAssociations\.%s\shell\CONVI"
    L = launcher(); exts = sorted(CTX_EXT)
    for e in exts:
        key = base % e
        _del_tree(winreg, winreg.HKEY_CURRENT_USER, key)
        if not on: continue
        k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key)
        winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, "Конвертировать в...")
        winreg.SetValueEx(k, "SubCommands", 0, winreg.REG_SZ, "")
        winreg.CloseKey(k)
        for n, f in enumerate(targets("x." + e)):
            kk = winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"{key}\\shell\\{n:02d}_{f}")
            winreg.SetValueEx(kk, "MUIVerb", 0, winreg.REG_SZ, f); winreg.CloseKey(kk)
            kc = winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"{key}\\shell\\{n:02d}_{f}\\command")
            winreg.SetValue(kc, "", winreg.REG_SZ, f'{L} --convert "%1" --to {f}'); winreg.CloseKey(kc)
        kk = winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"{key}\\shell\\99_batch")
        winreg.SetValueEx(kk, "MUIVerb", 0, winreg.REG_SZ, "Пакетная конвертация..."); winreg.CloseKey(kk)
        kc = winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"{key}\\shell\\99_batch\\command")
        winreg.SetValue(kc, "", winreg.REG_SZ, f'{L} --batch "%1"'); winreg.CloseKey(kc)

VK_MODS = {"Ctrl": (0x11, 0xA2, 0xA3), "Alt": (0x12, 0xA4, 0xA5), "Shift": (0x10, 0xA0, 0xA1)}
MOD_ORDER = ["Ctrl", "Alt", "Shift"]

def vk_token(vk):
    for n, g in VK_MODS.items():
        if vk in g: return n
    if 0x41 <= vk <= 0x5A or 0x30 <= vk <= 0x39: return chr(vk)
    if 0x70 <= vk <= 0x87: return "F%d" % (vk - 0x6F)
    if vk == 0x20: return "Space"
    return None

def token_vks(tok):
    if tok in VK_MODS: return VK_MODS[tok]
    if len(tok) == 1 and (tok.isdigit() or "A" <= tok <= "Z"): return (ord(tok),)
    if tok == "Space": return (0x20,)
    m = re.fullmatch(r"F(\d{1,2})", tok)
    if m and 1 <= int(m.group(1)) <= 24: return (0x6F + int(m.group(1)),)
    return None

def parse_combo(text):
    toks = [t.strip() for t in str(text).split("+") if t.strip()]
    if not toks or len(set(toks)) != len(toks): return None
    groups = []
    for t in toks:
        g = token_vks(t)
        if g is None: return None
        groups.append(g)
    return toks, groups

def combo_error(text):
    r = parse_combo(text)
    if not r: return "Неверное сочетание"
    toks = r[0]
    if not [t for t in toks if t not in VK_MODS]: return "Нужна обычная клавиша, а не только Ctrl / Alt / Shift"
    if len(toks) < 2 and not re.fullmatch(r"F\d+", toks[0]): return "Добавь Ctrl, Alt или Shift (одна клавиша допустима только F1-F24)"
    return None

def normalize_combo(toks):
    mods = [m for m in MOD_ORDER if m in toks]
    return "+".join(mods + [t for t in toks if t not in VK_MODS])

def combos_conflict(a, b):
    ra, rb = parse_combo(a), parse_combo(b)
    if not ra or not rb: return False
    sa, sb = set(ra[0]), set(rb[0])
    return sa <= sb or sb <= sa

def current_combos():
    return {"single": S.get("hk_single") or DEFAULTS["hk_single"], "batch": S.get("hk_batch") or DEFAULTS["hk_batch"]}

class ComboTracker:
    def __init__(self, combos=None, is_pressed=None):
        self.down = set(); self.eat = set(); self.is_pressed = is_pressed; self.combos = {}
        self.set_combos(combos or {"single": "Ctrl+F+H", "batch": "Ctrl+F+V"})
    def set_combos(self, combos):
        out = {}
        for act, text in combos.items():
            r = parse_combo(text)
            if r and not combo_error(text): out[act] = r[1]
        self.combos = out
    def _held(self, g, cur):
        if cur in g: return True
        if self.is_pressed: return bool(self.is_pressed(g[0]))
        return any(v in self.down for v in g)
    def feed(self, vk, is_down):
        if not is_down:
            self.down.discard(vk)
            if vk in self.eat: self.eat.discard(vk); return None, True
            return None, False
        again = vk in self.down; self.down.add(vk)
        if vk in self.eat: return None, True
        if again: return None, False
        for act, groups in self.combos.items():
            if any(vk in g for g in groups) and all(self._held(g, vk) for g in groups):
                sup = vk_token(vk) not in VK_MODS
                if sup: self.eat.add(vk)
                return act, sup
        return None, False

def kill_other_instances():
    if not IS_WIN: return
    ps = ("$me=%d; $pp=%d; Get-CimInstance Win32_Process | Where-Object { $_.ProcessId -ne $me -and $_.ProcessId -ne $pp -and $_.ProcessId -ne $PID -and "
          "($_.Name -in 'Convi.exe','AVI-CON Convert.exe' -or $_.CommandLine -like '*convi.py*' -or $_.CommandLine -like '*avicon.py*') } | "
          "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }") % (os.getpid(), os.getppid())
    try: subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], creationflags=NOWIN, timeout=20, capture_output=True)
    except Exception as e: log("Не удалось закрыть старые копии:", e)

def cleanup_legacy():
    try:
        old = Path(os.environ.get("APPDATA") or (Path.home() / ".config")) / "AVI-CON"
        DATA.mkdir(parents=True, exist_ok=True)
        if (old / "settings.json").exists() and not (DATA / "settings.json").exists(): shutil.copy(old / "settings.json", DATA / "settings.json")
    except Exception: pass
    if not IS_WIN: return
    try:
        import winreg
        for e in ALL_EXT: _del_tree(winreg, winreg.HKEY_CURRENT_USER, r"Software\Classes\SystemFileAssociations\.%s\shell\AVICON" % e)
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        try: winreg.DeleteValue(k, "AVI-CON Convert")
        except OSError: pass
        winreg.CloseKey(k)
    except Exception as e: log("Очистка старых записей:", e)

def apply_setting(k, v):
    if k not in DEFAULTS: return "Неизвестная настройка"
    if k == "mode" and v not in ("near", "folder", "replace"): return "Неверный режим сохранения"
    if k == "quality" and not (40 <= int(v) <= 100): return "Качество: от 40 до 100"
    if k == "dpi" and not (50 <= int(v) <= 600): return "DPI: от 50 до 600"
    if k in ("hk_single", "hk_batch"):
        e = combo_error(v)
        if e: return e
    err = None
    if k == "autostart":
        try: err = set_autostart(bool(v))
        except Exception as e: err = short_err(e)
    elif k == "ctx":
        try: err = set_ctx(bool(v))
        except Exception as e: err = short_err(e)
    if err: return err
    S[k] = v; save_settings()

def icon_image(size=256):
    from PIL import Image, ImageDraw
    import math
    S = 1024; im = Image.new("RGBA", (S, S), (0, 0, 0, 0)); d = ImageDraw.Draw(im); c = S / 2
    R = lambda r: (c - S * r, c - S * r, c + S * r, c + S * r)
    W = (255, 255, 255, 255); K = (12, 12, 12, 255)
    d.ellipse(R(.49), fill=K)
    d.pieslice(R(.43), -112.5, -67.5, fill=W)
    d.ellipse(R(.20), fill=K)
    for k in range(8):
        a = math.radians(-112.5 + 45 * k)
        d.line((c + S * .20 * math.cos(a), c + S * .20 * math.sin(a), c + S * .43 * math.cos(a), c + S * .43 * math.sin(a)), fill=W, width=int(S * .018))
    d.ellipse(R(.43), outline=W, width=int(S * .03))
    d.ellipse(R(.20), outline=W, width=int(S * .03))
    d.ellipse(R(.075), fill=W)
    return im.resize((size, size), Image.LANCZOS)

def make_ico(path):
    path = Path(path); path.parent.mkdir(parents=True, exist_ok=True)
    icon_image(256).save(path, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    return path

try:
    from PySide6.QtCore import Qt, QTimer, QPointF, QRectF, QObject, Signal, QPropertyAnimation, QEvent
    from PySide6.QtGui import (QPainter, QPainterPath, QColor, QPen, QBrush, QFont, QFontMetricsF, QIcon, QPixmap,
                               QCursor, QGuiApplication, QLinearGradient, QRadialGradient, QPalette)
    from PySide6.QtWidgets import (QApplication, QWidget, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                                   QCheckBox, QRadioButton, QFileDialog, QSystemTrayIcon, QMenu, QMessageBox,
                                   QSpinBox, QSlider, QLineEdit, QScrollArea, QFrame, QComboBox, QStackedWidget)
    from PySide6.QtNetwork import QLocalServer, QLocalSocket
    HAVE_QT = True
except Exception:
    HAVE_QT = False
    class _Stub:
        def __init__(self, *a, **k): pass
    Qt = QTimer = QPointF = QRectF = QPropertyAnimation = QEvent = _Stub
    QObject = QWidget = QDialog = _Stub
    Signal = lambda *a, **k: None

COL = dict(bg="#0b0b0b", panel="#111111", border="#282828", text="#f4f4f4", muted="#929292", dim="#5f5f5f",
           danger="#d7b4b4", sec_fill="#171717", sec_line="#303030", sec_hover="#292929", sec_hover_line="#5a5a5a",
           sec_empty="#101010", sec_empty_line="#222222")
W, H = 430, 480
CX, CY, RAD = 215, 236, 146
TAB_Y, BAR_H = 40, 38
BAR_Y = CY + RAD + 14

def qicon():
    pm = QPixmap(); bio = io.BytesIO(); icon_image(256).save(bio, "PNG"); pm.loadFromData(bio.getvalue(), "PNG")
    ic = QIcon()
    ic.addPixmap(pm)
    for s in (16, 24, 32, 48, 64, 128):
        q = QPixmap(); b = io.BytesIO(); icon_image(s).save(b, "PNG"); q.loadFromData(b.getvalue(), "PNG"); ic.addPixmap(q)
    return ic

def mkfont(pt, bold=False, spacing=0.0):
    f = QFont("Segoe UI"); f.setPointSizeF(pt)
    if bold: f.setWeight(QFont.Weight.DemiBold)
    if spacing: f.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, spacing)
    return f

def draw_icon(p, name, cx, cy, size, color, width=1.8, prog=1.0):
    p.save(); p.translate(cx, cy); s = size / 24.0; p.scale(s, s); p.translate(-12, -12)
    pen = QPen(QColor(color), width); pen.setCapStyle(Qt.PenCapStyle.RoundCap); pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    p.setPen(pen); p.setBrush(Qt.BrushStyle.NoBrush)
    L = lambda x0, y0, x1, y1: p.drawLine(QPointF(x0, y0), QPointF(x1, y1))
    def poly(*pts):
        path = QPainterPath(); path.moveTo(QPointF(*pts[0]))
        for q in pts[1:]: path.lineTo(QPointF(*q))
        p.drawPath(path)
    if name == "upload": L(12, 3, 12, 15); poly((7, 8), (12, 3), (17, 8)); poly((4, 15), (4, 20), (20, 20), (20, 15))
    elif name == "check":
        a, b, c = (5, 12.5), (10, 17.5), (19, 7); s1 = math.hypot(b[0] - a[0], b[1] - a[1]); s2 = math.hypot(c[0] - b[0], c[1] - b[1]); d = prog * (s1 + s2)
        if d > 0.05:
            if d <= s1: f = d / s1; poly(a, (a[0] + (b[0] - a[0]) * f, a[1] + (b[1] - a[1]) * f))
            else: f = (d - s1) / s2; poly(a, b, (b[0] + (c[0] - b[0]) * f, b[1] + (c[1] - b[1]) * f))
    elif name == "open": poly((14, 4), (20, 4), (20, 10)); L(20, 4, 11, 13); poly((18, 14), (18, 19), (5, 19), (5, 6), (10, 6))
    elif name == "x": L(6, 6, 18, 18); L(18, 6, 6, 18)
    elif name == "settings":
        L(20, 7, 11, 7); L(14, 17, 5, 17); p.drawEllipse(QPointF(17, 17), 3, 3); p.drawEllipse(QPointF(7, 7), 3, 3)
    elif name == "folder": poly((3, 6), (9, 6), (11, 8), (21, 8), (21, 19), (3, 19), (3, 6))
    elif name == "image":
        p.drawRoundedRect(QRectF(3, 3, 18, 18), 2, 2); p.drawEllipse(QPointF(9, 9), 2, 2); poly((21, 15), (16, 10), (6, 21))
    elif name == "film":
        p.drawRoundedRect(QRectF(2, 2, 20, 20), 2, 2); L(7, 2, 7, 22); L(17, 2, 17, 22); L(2, 12, 22, 12)
    elif name == "music":
        p.drawEllipse(QPointF(6, 18), 3, 3); p.drawEllipse(QPointF(18, 16), 3, 3); poly((9, 18), (9, 5), (21, 3), (21, 16))
    elif name == "doc":
        p.drawRoundedRect(QRectF(5, 3, 14, 18), 2, 2); L(8, 9, 16, 9); L(8, 13, 16, 13); L(8, 17, 13, 17)
    elif name == "table": p.drawRoundedRect(QRectF(3, 4, 18, 16), 2, 2); L(3, 10, 21, 10); L(3, 15, 21, 15); L(9, 4, 9, 20)
    elif name == "pres": p.drawRoundedRect(QRectF(3, 4, 18, 12), 2, 2); L(12, 16, 12, 20); L(8, 20, 16, 20); poly((7, 12), (10, 9), (13, 11), (17, 7))
    elif name == "layers":
        poly((12, 2), (22, 7), (12, 12), (2, 7), (12, 2)); poly((2, 12), (12, 17), (22, 12)); poly((2, 17), (12, 22), (22, 17))
    p.restore()

def clamp01(x): return 0.0 if x < 0 else 1.0 if x > 1 else x
def seg(t, a, b): return clamp01((t - a) / (b - a))
def ease_out_cubic(x): return 1 - (1 - x) ** 3
def ease_out_back(x):
    c1 = 1.5; c3 = c1 + 1; return 1 + c3 * (x - 1) ** 3 + c1 * (x - 1) ** 2
def anim_values(t):
    return dict(cs=ease_out_back(seg(t, 0.00, 0.30)), ca=seg(t, 0.00, 0.18), rs=ease_out_cubic(seg(t, 0.22, 0.66)),
                ls=seg(t, 0.58, 0.88), ts=ease_out_cubic(seg(t, 0.72, 1.00)))

def ftype(f): return f.get("type") or type_of(f["name"])

def mix_hex(a, b, t):
    t = clamp01(t); ca = [int(a[i:i + 2], 16) for i in (1, 3, 5)]; cb = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(ca, cb))

FMT_HINT = {"JPG": "универсальное фото, подходит всем", "PNG": "без потерь, умеет прозрачность", "WEBP": "компактно, для сайтов и мессенджеров",
    "AVIF": "самое компактное фото нового поколения", "BMP": "без сжатия, большие файлы", "GIF": "короткая анимация", "TIFF": "печать и архивы без потерь",
    "ICO": "значок для программ и сайтов", "PDF": "документ, одинаково выглядит везде", "MP4": "видео, которое играет везде", "AVI": "старый формат видео",
    "MKV": "много дорожек и субтитров", "MOV": "видео Apple", "WEBM": "видео для сайтов", "WMV": "видео Windows", "MP3": "музыка, играет везде",
    "WAV": "без сжатия, студийное качество", "OGG": "открытый формат аудио", "FLAC": "без потерь, для ценителей", "M4A": "аудио Apple и YouTube",
    "AAC": "хорошее качество при малом размере", "OPUS": "лучший голос при малом размере", "WMA": "аудио Windows", "DOCX": "документ Word",
    "TXT": "простой текст", "HTML": "веб-страница", "MD": "Markdown для заметок и GitHub", "RTF": "текст с оформлением", "XLSX": "таблица Excel",
    "CSV": "таблица, читает любая программа", "JSON": "данные для программистов", "PPTX": "презентация PowerPoint"}
TOOL_HINT = {"Сжать": "уменьшить размер файла", "Размер": "изменить ширину и высоту", "Поворот": "повернуть или отразить", "Ч/Б": "сделать чёрно-белым",
    "Обрезка": "обрезать до нужного формата кадра", "Подпись": "добавить текст-водяной знак", "Без EXIF": "убрать геометки и данные о камере",
    "В PDF": "собрать фото в один PDF", "Склеить": "объединить PDF в один файл", "Разбить": "разделить PDF на части",
    "Обрезать": "оставить нужный фрагмент по времени", "Без звука": "убрать звуковую дорожку"}

def open_path(p):
    try:
        if IS_WIN: os.startfile(str(p))
        elif sys.platform == "darwin": subprocess.Popen(["open", str(p)])
        else: subprocess.Popen(["xdg-open", str(p)])
    except Exception: pass

def apply_dark(app):
    app.setStyle("Fusion"); pal = QPalette(); R = QPalette.ColorRole
    for role, col in ((R.Window, "#0b0b0b"), (R.WindowText, "#f4f4f4"), (R.Base, "#141414"), (R.AlternateBase, "#1a1a1a"), (R.ToolTipBase, "#151515"),
                      (R.ToolTipText, "#f4f4f4"), (R.Text, "#f4f4f4"), (R.Button, "#171717"), (R.ButtonText, "#eaeaea"), (R.BrightText, "#ffffff"),
                      (R.Link, "#8ab4ff"), (R.Highlight, "#f4f4f4"), (R.HighlightedText, "#080808"), (R.PlaceholderText, "#6a6a6a")): pal.setColor(role, QColor(col))
    for role in (R.Text, R.ButtonText, R.WindowText): pal.setColor(QPalette.ColorGroup.Disabled, role, QColor("#5e5e5e"))
    app.setPalette(pal)

TOOL_NAMES = {"image": "Фото", "video": "Видео", "audio": "Аудио", "pdf": "PDF"}
MERGE_OPS = ("images_pdf", "pdf_merge")
def tool_group(f):
    if ext_of(f["name"]) == "pdf": return "pdf"
    t = ftype(f); return t if t in ("image", "video", "audio") else None

def tools_for(group, n):
    if group == "image": return [("Сжать", "compress"), ("Размер", "resize"), ("Поворот", "rotate"), ("Ч/Б", "bw"), ("Обрезка", "crop"), ("Подпись", "watermark"), ("Без EXIF", "strip")] + ([("В PDF", "images_pdf")] if n > 1 else [])
    if group == "video": return [("Сжать", "vid_compress"), ("Обрезать", "trim"), ("Без звука", "vid_mute")]
    if group == "audio": return [("Обрезать", "trim")]
    return ([("Склеить", "pdf_merge")] if n > 1 else []) + [("Разбить", "pdf_split"), ("Поворот", "pdf_rotate"), ("Сжать", "pdf_compress")]

def fmt_size(b):
    return f"{b} Б" if b < 1024 else f"{b/1024:.0f} КБ" if b < 1048576 else f"{b/1048576:.1f} МБ" if b < 1073741824 else f"{b/1073741824:.2f} ГБ"

def plural(n):
    return "файл" if n % 10 == 1 and n % 100 != 11 else "файла" if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14) else "файлов"

def save_text():
    m = S.get("mode")
    if m == "folder": return "в папку " + Path(S.get("outdir_path") or default_outdir()).name
    return "рядом с оригиналом" if m != "replace" else ("вместо оригинала, он в корзину" if S.get("trash") else "вместо оригинала")

def reveal(path):
    try:
        if IS_WIN: subprocess.Popen(["explorer", "/select," + str(path)])
        elif sys.platform == "darwin": subprocess.Popen(["open", "-R", str(path)])
        else: subprocess.Popen(["xdg-open", str(Path(path).parent)])
    except Exception: pass

class RoundWindow(QWidget):
    def __init__(self, ctrl):
        super().__init__()
        self.ctrl = ctrl
        self.setWindowTitle(APP); self.setWindowIcon(qicon())
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setAcceptDrops(True); self.setMouseTracking(True); self.setFixedSize(W, H)
        self.mode = "single"; self.files = []; self.sel = {}; self.tab = None
        self.state = "idle"; self.msg = ""; self.flash = ""; self.target = None; self.last_out = None
        self.job = None; self.pct = 0.0; self.shown = 0.0; self.jview = {}; self.title = ""
        self.over = False; self.hover = None; self.press = None; self.press_pos = None
        self.hits = []; self.dialog_open = False; self._prefix = False; self._fade = None
        self.t = 1.0; self.A = anim_values(1.0); self.anim_dir = 0; self._after = "hide"; self._fx_last = 0.0; self._shown_at = 0.0
        self._fx = QTimer(self); self._fx.setInterval(15); self._fx.timeout.connect(self._fx_tick)
        self.poll = QTimer(self); self.poll.setInterval(120); self.poll.timeout.connect(self._poll_job)
        self.anim = QTimer(self); self.anim.setInterval(33); self.anim.timeout.connect(self._tick)
        self.flash_t = QTimer(self); self.flash_t.setSingleShot(True); self.flash_t.timeout.connect(self._clear_flash)
        self.prefix_t = QTimer(self); self.prefix_t.setSingleShot(True); self.prefix_t.timeout.connect(lambda: setattr(self, "_prefix", False))
        self.hv = {}; self.done_t = 1.0; self.thumb = None; self._thumb_png = None; self._tkey = None; self._vt = time.perf_counter()
        self.vis = QTimer(self); self.vis.setInterval(33); self.vis.timeout.connect(self._vis_tick)

    def busy(self): return self.state == "running"

    def set_mode(self, mode):
        if self.busy(): return
        self.mode = mode; self.reset();

    def showEvent(self, e): self._vt = time.perf_counter(); self.vis.start(); super().showEvent(e)

    def hideEvent(self, e): self.vis.stop(); super().hideEvent(e)

    def _vis_tick(self):
        now = time.perf_counter(); dt = min(.1, now - self._vt); self._vt = now; need = False
        fm = self.ring_formats()[0]; hv = self.hover
        for i in range(max(len(fm), 8)):
            tgt = 1.0 if (hv and hv[0] == "fmt" and hv[1] == i) else 0.0; cur = self.hv.get(i, 0.0)
            if abs(tgt - cur) > 0.004: self.hv[i] = cur + (tgt - cur) * min(1.0, dt * 14); need = True
            elif cur != tgt: self.hv[i] = tgt
        if self.done_t < 1.0: self.done_t = min(1.0, self.done_t + dt / 0.9); need = True
        b = self._thumb_png
        if b: self._thumb_png = None; pm = QPixmap(); pm.loadFromData(b, "PNG"); self.thumb = pm; need = True
        if self.over and self.A["rs"] > 0.99: need = True
        if need: self.update()

    def load_thumb(self, f):
        self.thumb = None; self._thumb_png = None; self._tkey = f["path"]; t = ftype(f)
        if t not in ("image", "video"): return
        def work():
            b = make_thumb(f["path"], t, 220)
            if self._tkey == f["path"]: self._thumb_png = b
        threading.Thread(target=work, daemon=True).start()

    def cur_type(self):
        if self.mode == "single" and self.files: return ftype(self.files[0])
        return self.tab if self.mode == "batch" else None

    def last_fmt(self):
        t = self.cur_type(); return (S.get("last") or {}).get(t) if t else None

    def remember(self, t, fmt):
        if t and fmt: S.setdefault("last", {})[t] = fmt; save_settings()

    def hint(self):
        fm = self.ring_formats()[0]; i = self.hover[1]
        if not (0 <= i < len(fm)): return ""
        lab = fm[i]; h = (TOOL_HINT if self.mode == "tools" else FMT_HINT).get(lab, "")
        return (f"{lab} - {h}" if h else lab) + (" · последний выбор" if self.mode != "tools" and lab == self.last_fmt() else "")

    def paste(self):
        if self.busy(): return
        cb = QGuiApplication.clipboard(); md = cb.mimeData()
        if md is None: return
        if md.hasUrls():
            fs = [x for x in (stat_path(u.toLocalFile()) for u in md.urls() if u.isLocalFile()) if x]
            if fs: self.take_files(fs); return
        if md.hasImage():
            d = default_outdir(); d.mkdir(parents=True, exist_ok=True); path = unique(d / ("Скриншот " + time.strftime("%Y-%m-%d %H-%M-%S") + ".png"))
            if cb.image().save(str(path), "PNG"): self.take_files([stat_path(path)]); return
        self.notice("В буфере нет файлов или картинки")

    def reset(self):
        self.thumb = None; self._tkey = None; self._thumb_png = None
        self.files = []; self.sel = {}; self.tab = None; self.state = "idle"; self.msg = ""; self.target = None; self.job = None
        self.poll.stop(); self.update()

    def notice(self, text, ms=2800):
        self.flash = text; self.flash_t.start(ms); self.update()

    def _clear_flash(self): self.flash = ""; self.update()

    def ring_formats(self):
        if self.mode == "single" and self.files and self.state in ("chosen", "running"):
            return targets(self.files[0]["name"], ftype(self.files[0])), (self.target if self.state == "running" else None), self.state == "running"
        if self.mode == "batch" and self.state == "batch" and self.tab:
            return list(TYPES[self.tab]["to"]), self.sel.get(self.tab), False
        if self.mode == "tools" and self.state == "tools" and self.tab:
            return [l for l, _ in tools_for(self.tab, sum(1 for f in self.files if tool_group(f) == self.tab))], None, False
        return [], None, False

    def take_files(self, fs):
        if self.busy() or not fs: return
        if self.mode == "single":
            if len(fs) > 1: self.notice("Для нескольких файлов - режим «Несколько»")
            self.pick_single(fs[0])
        elif self.mode == "tools": self.add_tools(fs)
        else: self.add_batch(fs)

    def add_tools(self, fs):
        for f in fs:
            if tool_group(f) and not any(b["path"] == f["path"] for b in self.files): self.files.append(f)
        groups = [g for g in ("image", "video", "audio", "pdf") if any(tool_group(f) == g for f in self.files)]
        if not groups: self.files = []; self.state = "idle"; self.notice("Инструменты есть для фото, видео, аудио и PDF"); return
        if self.tab not in groups: self.tab = groups[0]
        self.state = "tools"; self.update()

    def run_tool(self, idx):
        n = sum(1 for f in self.files if tool_group(f) == self.tab); lst = tools_for(self.tab, n)
        if not (0 <= idx < len(lst)): return
        label, op = lst[idx]; files = [f for f in self.files if tool_group(f) == self.tab]
        self.dialog_open = True
        try: params = ask_tool_params(op, self)
        finally: self.dialog_open = False
        if params is None: return
        if op in MERGE_OPS: items = [dict(op=op, paths=[f["path"] for f in files], params=params)]
        else: items = [dict(op=op, path=f["path"], params=params) for f in files]
        self._start(items, label)

    def pick_single(self, f, to=None):
        t = ftype(f)
        if not t:
            self.files = []; self.state = "error"; self.msg = "Формат не поддерживается: ." + ext_of(f["name"]); self.update(); return
        self.files = [f]; self.state = "chosen"; self.target = None; self.load_thumb(f); self.update()
        if to and to.upper() in targets(f["name"], ftype(f)):
            idx = targets(f["name"], ftype(f)).index(to.upper())
            QTimer.singleShot(450, lambda: self.run_single(idx) if self.state == "chosen" and self.files and self.files[0] is f else None)

    def add_batch(self, fs):
        for f in fs:
            if not any(b["path"] == f["path"] for b in self.files): self.files.append(f)
        types = []
        for f in self.files:
            t = ftype(f)
            if t and t not in types: types.append(t)
        self.sel = {t: self.sel.get(t) for t in types}
        if not self.tab or self.tab not in types: self.tab = next((t for t in types if not self.sel[t]), types[0] if types else None)
        if not types:
            self.files = []; self.state = "idle"; self.notice("Нет поддерживаемых файлов"); return
        self.state = "batch"; self.update()

    def choose_dialog(self):
        pat = lambda xs: " ".join("*." + e for e in sorted(xs))
        flt = ";;".join(["Все поддерживаемые (" + pat(ALL_EXT) + ")"] + [f"{v['label']} (" + pat(v["ext"]) + ")" for v in TYPES.values()] + ["Любые файлы (*.*)"])
        self.dialog_open = True
        try:
            if self.mode == "batch": paths, _ = QFileDialog.getOpenFileNames(self, "Выбери файлы", "", flt)
            else:
                one, _ = QFileDialog.getOpenFileName(self, "Выбери файл", "", flt); paths = [one] if one else []
        finally: self.dialog_open = False
        self.take_files([x for x in (stat_path(q) for q in paths) if x])

    def _start(self, items, target_label):
        job, err = start_job(items)
        if not job: self.state = "error"; self.msg = err or "Ошибка"; self.update(); return
        self.job = job; self.pct = 0.0; self.shown = 0.0; self.jview = {}; self.target = target_label; self.state = "running"
        self.poll.start(); self.anim.start(); self.update()

    def run_single(self, idx):
        fmts, _, _ = self.ring_formats()
        if not (0 <= idx < len(fmts)) or not self.files: return
        f = self.files[0]; self.remember(ftype(f), fmts[idx])
        self._start([dict(path=f["path"], to=fmts[idx])], fmts[idx])

    def start_batch(self):
        items = [dict(path=f["path"], to=self.sel[ftype(f)]) for f in self.files if ftype(f) and self.sel.get(ftype(f))]
        pics = [it for it in items if it["to"] == "PDF" and ext_of(it["path"]) != "pdf" and any(f["path"] == it["path"] and ftype(f) == "image" for f in self.files)]
        if S.get("merge_pdf", True) and len(pics) > 1:
            items = [it for it in items if it not in pics] + [dict(op="images_pdf", paths=[it["path"] for it in pics], params={})]
        for t, v in self.sel.items(): self.remember(t, v)
        if items: self._start(items, "")

    def cancel(self):
        if self.job: self.job.kill()
        self.poll.stop(); self.anim.stop(); self.reset()

    def _tick(self):
        self.shown += (self.pct - self.shown) * 0.3
        if not self.busy(): self.anim.stop()
        self.update()

    def _poll_job(self):
        if not self.job: self.poll.stop(); return
        v = self.job.view(); self.jview = v; self.pct = v["pct"]
        if v["state"] != "run": self.poll.stop(); self.anim.stop(); self.finish(v)
        self.update()

    def finish(self, v):
        if v["state"] == "cancel": self.reset(); return
        ok = [r for r in v["results"] if r.get("ok")]; bad = [r for r in v["results"] if not r.get("ok")]
        if not ok and bad: self.state = "error"; self.msg = bad[0].get("error") or "Ошибка"; self.ctrl.notify("Ошибка: " + self.msg); return
        self.last_out = ok[-1]["dst"] if ok else None
        self.state = "done"; self.done_t = 0.0
        si = sum(r.get("size_in", 0) for r in ok); so = sum(r.get("size_out", 0) for r in ok if "size_out" in r)
        gain = f" · {fmt_size(si)} → {fmt_size(so)}" + (f" (на {100 - 100 * so // si}% меньше)" if si and 0 < so < si and len(ok) == 1 else "") if si and so else ""
        self.msg = (f"ошибок: {len(bad)} · {bad[0].get('error')}" if bad else f"Сохранён {save_text()}{gain}")
        self.title = Path(self.last_out).name if (len(v["results"]) == 1 and self.last_out) else f"Готово: {len(ok)}"
        self.ctrl.notify("Готово: " + (Path(self.last_out).name if len(ok) == 1 and self.last_out else f"{len(ok)} {plural(len(ok))}"))

    def open_request(self, mode, files=None, to=None, pre=None):
        self.popup(at_cursor=not self.isVisible())
        if self.busy(): return
        self.mode = mode; self.reset()
        files = files or []
        if not files: return
        if mode == "single": self.pick_single(files[0], to)
        else:
            self.add_batch(files)
            if pre and self.state == "batch":
                for t in list(self.sel):
                    if pre.get(t): self.sel[t] = pre[t]
                self.update()
                if all(self.sel.values()): QTimer.singleShot(500, self.start_batch)

    def popup(self, at_cursor=True):
        if at_cursor:
            pos = QCursor.pos(); scr = QGuiApplication.screenAt(pos) or QGuiApplication.primaryScreen(); g = scr.availableGeometry()
            self.move(int(min(max(g.left(), pos.x() - W // 2), g.right() - W)), int(min(max(g.top(), pos.y() - CY), g.bottom() - H)))
        if not self.isVisible():
            self.t = 0.0; self.A = anim_values(0.0); self.show()
            self.raise_(); self.activateWindow(); self.animate(1)
        elif self.anim_dir < 0 or self.t < 1.0:
            self.raise_(); self.activateWindow(); self.animate(1)
        else:
            self.raise_(); self.activateWindow()

    def animate(self, direction, after="hide"):
        log("анимация:", "показ" if direction > 0 else "скрытие", f"t={self.t:.2f}")
        self.anim_dir = direction; self._after = after; self._fx_last = time.perf_counter()
        if direction > 0: self._shown_at = self._fx_last
        if not self._fx.isActive(): self._fx.start()

    def _fx_tick(self):
        try:
            now = time.perf_counter(); dt = max(0.0, now - self._fx_last); self._fx_last = now
            if self.anim_dir == 0: self._fx.stop(); return
            self.t = clamp01(self.t + self.anim_dir * dt / (0.85 if self.anim_dir > 0 else 0.55))
            self.A = anim_values(self.t)
            if self.anim_dir > 0 and self.t >= 1.0: self.anim_dir = 0; self._fx.stop()
            elif self.anim_dir < 0 and self.t <= 0.0:
                self.anim_dir = 0; self._fx.stop(); self.t = 0.0
                if self._after == "quit": self.ctrl.quit()
                else: self.hide()
            self.update()
        except Exception:
            log("Ошибка анимации:\n" + traceback.format_exc()); self.finish_anim()

    def finish_anim(self):
        self._fx.stop(); self.anim_dir = 0; self.t = 1.0; self.A = anim_values(1.0); self.update()

    def shown_ok(self):
        return self.isVisible() and self.anim_dir >= 0 and self.t > 0.05

    def center_on_screen(self):
        g = QGuiApplication.primaryScreen().availableGeometry()
        self.move(g.left() + (g.width() - W) // 2, g.top() + (g.height() - H) // 2)

    def geom(self):
        outer = RAD - 4; inner = outer * 0.52
        return QPointF(CX, CY), outer, inner, inner - 8

    @staticmethod
    def pt(c, r, deg):
        a = math.radians(deg); return QPointF(c.x() + math.sin(a) * r, c.y() - math.cos(a) * r)

    def sector(self, c, outer, inner, a0, a1):
        path = QPainterPath(); path.moveTo(self.pt(c, outer, a0))
        path.arcTo(QRectF(c.x() - outer, c.y() - outer, outer * 2, outer * 2), 90 - a0, -(a1 - a0))
        path.lineTo(self.pt(c, inner, a1))
        path.arcTo(QRectF(c.x() - inner, c.y() - inner, inner * 2, inner * 2), 90 - a1, (a1 - a0))
        path.closeSubpath(); return path

    def index_at(self, x, y):
        fmts, _, _ = self.ring_formats(); c, outer, inner, _ = self.geom()
        d = math.hypot(x - c.x(), y - c.y())
        if not fmts or d < inner or d > outer + 3: return -1
        ang = (math.degrees(math.atan2(x - c.x(), -(y - c.y()))) + 360) % 360; step = 360 / len(fmts)
        i = int(((ang + step / 2) % 360) / step); return i if i < len(fmts) else -1

    def hit_at(self, x, y):
        if self.t < 1.0: return None
        for r, act in reversed(self.hits):
            if r.contains(QPointF(x, y)): return act
        c, outer, inner, cr = self.geom(); d = math.hypot(x - c.x(), y - c.y())
        if d <= cr: return ("center", None)
        i = self.index_at(x, y)
        if i >= 0: return ("fmt", i)
        return None

    def dtext(self, p, text, cx, cy, font, color, maxw=None):
        p.setFont(font); p.setPen(QColor(color)); fm = QFontMetricsF(font)
        if maxw: text = fm.elidedText(text, Qt.TextElideMode.ElideMiddle, maxw)
        w = fm.horizontalAdvance(text) + 6; h = fm.height()
        p.drawText(QRectF(cx - w / 2, cy - h / 2, w, h), Qt.AlignmentFlag.AlignCenter, text)

    def pill(self, p, r, fill, border, rad=None):
        p.setPen(QPen(QColor(border), 1)); p.setBrush(QBrush(QColor(fill)))
        rad = r.height() / 2 if rad is None else rad; p.drawRoundedRect(r, rad, rad)

    def paintEvent(self, ev):
        p = QPainter(self); self.hits = []
        try:
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True); p.setRenderHint(QPainter.RenderHint.TextAntialiasing, True); p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
            self.paint_tabs(p); self.paint_ring(p); self.paint_center(p); self.paint_bar(p)
        except Exception:
            log("Ошибка отрисовки:\n" + traceback.format_exc()); self.finish_anim()
            try:
                p.setPen(QPen(QColor("#ffffff"), 2)); p.setBrush(QBrush(QColor("#0d0d0d"))); p.drawEllipse(QPointF(CX, CY), RAD, RAD)
                p.drawText(QPointF(CX - 90, CY), "Ошибка отрисовки - см. convi.log")
            except Exception: pass
        finally: p.end()

    def paint_tabs(self, p):
        ts = self.A["ts"]
        if ts <= 0.001: return
        p.save(); p.setOpacity(ts); p.translate(0, (ts - 1) * 12)
        labels = [("single", "ОДИН ФАЙЛ", 108), ("batch", "НЕСКОЛЬКО", 100), ("tools", "ИНСТРУМЕНТЫ", 118)]; gap = 6
        x = (W - sum(w for _, _, w in labels) - gap) / 2; f = mkfont(7.5, True, 1.3)
        for key, text, w in labels:
            r = QRectF(x, TAB_Y, w, 26); on = self.mode == key; hov = self.hover == ("tab", key)
            self.pill(p, r, "#f4f4f4" if on else ("#161616" if hov else "#0c0c0c"), "#f4f4f4" if on else "#2c2c2c")
            self.dtext(p, text, r.center().x(), r.center().y(), f, "#080808" if on else ("#ffffff" if hov else "#9a9a9a"))
            self.hits.append((r, ("tab", key))); x += w + gap
        p.restore()

    def paint_ring(self, p):
        A = self.A; rs = A["rs"]
        if rs <= 0.001: return
        c, outer, inner, cr = self.geom(); fmts, selected, dim = self.ring_formats()
        base = min(1.0, rs * 1.5); o = inner + (outer - inner) * rs; rd = cr + (RAD - cr) * rs
        p.save(); p.setOpacity(base)
        p.translate(c.x(), c.y()); p.rotate(-38 * (1 - rs)); p.translate(-c.x(), -c.y())
        gd = QRadialGradient(QPointF(c.x(), c.y() - 40), RAD * 1.25); gd.setColorAt(0, QColor("#191919")); gd.setColorAt(1, QColor("#080808"))
        p.setPen(QPen(QColor("#2a2a2a" if not self.over else "#ffffff"), 1.4 if self.over else 1)); p.setBrush(QBrush(gd))
        p.drawEllipse(c, rd, rd)
        count = len(fmts); drawn = count or 8; step = 360 / drawn
        hov_i = self.hover[1] if self.hover and self.hover[0] == "fmt" else -1
        labels = []
        for i in range(drawn):
            a0 = -step / 2 + i * step + 1.4; a1 = -step / 2 + (i + 1) * step - 1.4
            fmt = fmts[i] if i < count else ""
            if not fmt: fill, line = COL["sec_empty"], ("#4a4a4a" if self.over else COL["sec_empty_line"])
            elif fmt == selected: fill, line = "#f4f4f4", "#ffffff"
            elif dim: fill, line = "#0f0f0f", "#222222"
            else:
                k = self.hv.get(i, 0.0); fill = mix_hex(COL["sec_fill"], COL["sec_hover"], k); line = mix_hex(COL["sec_line"], COL["sec_hover_line"], k)
            if fmt and self.mode == "batch" and self.tab and self.sel.get(self.tab) == fmt: fill, line = "#f4f4f4", "#ffffff"
            p.setPen(QPen(QColor(line), 1.0)); p.setBrush(QBrush(QColor(fill))); p.drawPath(self.sector(c, o, inner, a0, a1))
            if fmt: labels.append((i, fmt, self.pt(c, (o + inner) / 2, (a0 + a1) / 2), fill == "#f4f4f4", self.pt(c, inner + 9, (a0 + a1) / 2)))
        lastf = self.last_fmt() if self.mode != "tools" else None
        for i, fmt, pos, lit, dpos in labels:
            la = clamp01(A["ls"] * 1.6 - i * 0.07)
            if la <= 0.001: continue
            p.setOpacity(la)
            self.dtext(p, fmt, pos.x(), pos.y(), mkfont(9.5, True), "#090909" if lit else ("#555555" if dim else "#ededed"))
            if fmt == lastf and not dim:
                p.setPen(Qt.PenStyle.NoPen); p.setBrush(QBrush(QColor("#0a0a0a" if lit else "#e8e8e8"))); p.drawEllipse(dpos, 2.3, 2.3)
        p.restore()
        if rs > 0.99:
            now = time.perf_counter()
            if self.over:
                ph = (now * 1.1) % 1.0; p.setBrush(Qt.BrushStyle.NoBrush); p.setPen(QPen(QColor(255, 255, 255, int(150 * (1 - ph))), 1.6)); p.drawEllipse(c, RAD + 3 + ph * 16, RAD + 3 + ph * 16)

    def paint_center(self, p):
        c, outer, inner, cr = self.geom(); st = self.state
        hov = self.hover == ("center", None)
        A = self.A
        if A["ca"] <= 0.001 and A["cs"] <= 0.001: return
        p.save(); p.setOpacity(A["ca"])
        p.setPen(QPen(QColor("#ffffff" if self.over else ("#4a4a4a" if hov else "#2b2b2b")), 1.5 if self.over else 1))
        p.setBrush(QBrush(QColor("#151515" if (self.over or hov) and not self.busy() else "#0b0b0b"))); p.drawEllipse(c, cr * A["cs"], cr * A["cs"])
        p.restore()
        if A["ls"] <= 0.001: return
        p.save(); p.setOpacity(A["ls"])
        cx, cy = c.x(), c.y(); mw = cr * 2 - 22
        title = mkfont(11, True); sub = mkfont(7.5); big = mkfont(24, True)
        if st == "idle":
            draw_icon(p, "upload", cx, cy - 26, 22, "#8a8a8a")
            self.dtext(p, "Брось файл" if self.mode == "single" else "Брось файлы", cx, cy + 4, title, "#f5f5f5")
            self.dtext(p, "или нажми", cx, cy + 22, sub, "#777777")
        elif st == "chosen":
            f = self.files[0]; t = ftype(f)
            if self.thumb is not None:
                box = QRectF(cx - cr, cy - cr, cr * 2, cr * 2); p.save(); clip = QPainterPath(); clip.addEllipse(c, cr - 1.5, cr - 1.5); p.setClipPath(clip)
                p.drawPixmap(box, self.thumb, QRectF(0, 0, self.thumb.width(), self.thumb.height()))
                g = QLinearGradient(0, cy - cr, 0, cy + cr); g.setColorAt(0, QColor(0, 0, 0, 30)); g.setColorAt(0.5, QColor(0, 0, 0, 70)); g.setColorAt(1, QColor(0, 0, 0, 225))
                p.fillRect(box, QBrush(g)); p.restore()
                self.dtext(p, f["name"], cx, cy + 16, mkfont(9, True), "#ffffff", mw)
                self.dtext(p, f"{ext_of(f['name']).upper()} · {fmt_size(f['size'])}", cx, cy + 33, sub, "#d0d0d0")
            else:
                draw_icon(p, TYPES[t]["icon"], cx, cy - 32, 20, "#8a8a8a")
                self.dtext(p, f["name"], cx, cy - 6, mkfont(10, True), "#f5f5f5", mw)
                self.dtext(p, f"{ext_of(f['name']).upper()} · {fmt_size(f['size'])}", cx, cy + 14, sub, "#888888")
        elif st == "running":
            n = self.jview.get("total", 1) or 1
            if self.mode == "batch" or n > 1: self.dtext(p, f"{min(self.jview.get('idx', 0) + 1, n)}/{n}", cx, cy - 4, big, "#f5f5f5")
            else: self.dtext(p, f"{int(round(self.shown * 100))}%", cx, cy - 4, big, "#f5f5f5")
            self.dtext(p, f"в {self.target}" if self.target else "конвертация", cx, cy + 22, sub, "#888888")
            rr = cr + 7; rect = QRectF(cx - rr, cy - rr, rr * 2, rr * 2)
            p.setBrush(Qt.BrushStyle.NoBrush); p.setPen(QPen(QColor("#222222"), 2)); p.drawEllipse(c, rr, rr)
            pen = QPen(QColor("#ffffff"), 2.6); pen.setCapStyle(Qt.PenCapStyle.RoundCap); p.setPen(pen)
            p.drawArc(rect, 90 * 16, int(-max(0.01, min(1.0, self.shown)) * 360 * 16))
        elif st == "done":
            k = self.done_t
            if k < 1.0:
                rr = cr + 4 + 46 * ease_out_cubic(k); p.setBrush(Qt.BrushStyle.NoBrush); p.setPen(QPen(QColor(255, 255, 255, int(170 * (1 - k))), 2)); p.drawEllipse(c, rr, rr)
            draw_icon(p, "check", cx, cy - 24, 30, "#ffffff", 2.0, clamp01(k * 2.0))
            self.dtext(p, getattr(self, "title", "Готово"), cx, cy + 8, mkfont(10, True), "#f5f5f5", mw)
            self.dtext(p, "нажми для нового", cx, cy + 26, sub, "#777777")
        elif st == "error":
            draw_icon(p, "x", cx, cy - 24, 26, COL["danger"], 2.0)
            self.dtext(p, "Ошибка", cx, cy + 6, title, "#f5f5f5")
            self.dtext(p, "нажми, чтобы заново", cx, cy + 24, sub, "#777777")
        elif st == "tools":
            n = sum(1 for f in self.files if tool_group(f) == self.tab)
            draw_icon(p, "layers", cx, cy - 30, 20, "#8a8a8a")
            self.dtext(p, f"{n} {plural(n)}", cx, cy - 4, title, "#f5f5f5"); self.dtext(p, "выбери инструмент", cx, cy + 16, sub, "#888888")
        elif st == "batch":
            n = len(self.files); ready = bool(self.sel) and all(self.sel.values())
            if ready:
                self.dtext(p, str(n), cx, cy - 22, mkfont(26, True), "#f5f5f5"); self.dtext(p, f"{plural(n)} готово", cx, cy + 2, sub, "#888888")
                r = QRectF(cx - 36, cy + 14, 72, 24); self.pill(p, r, "#ffffff", "#ffffff")
                self.dtext(p, "СТАРТ", r.center().x(), r.center().y(), mkfont(8, True, 1.0), "#000000"); self.hits.append((r, ("go", None)))
            else:
                draw_icon(p, "layers", cx, cy - 30, 20, "#8a8a8a")
                self.dtext(p, f"{n} {plural(n)}", cx, cy - 4, title, "#f5f5f5")
                self.dtext(p, ("формат: " + TYPES[self.tab]["label"].lower()) if self.tab else "брось ещё", cx, cy + 16, sub, "#888888")
        p.restore()

    def paint_bar(self, p):
        ts = self.A["ts"]
        if ts <= 0.001: return
        p.save(); p.setOpacity(ts); p.translate(0, (1 - ts) * 14)
        st = self.state; btns = []
        if st == "running": btns = ["cancel"]
        else:
            if st == "done" and self.last_out: btns += ["open", "reveal"]
            btns += ["settings", "close"]
        chips = []; segs = []
        hint = self.hint() if (self.hover and self.hover[0] == "fmt" and not self.busy() and self.t >= 1.0) else ""
        if hint: segs = [(hint, "#f4f4f4")]
        elif st == "batch" and not self.busy():
            for t in self.sel:
                cnt = sum(1 for f in self.files if ftype(f) == t)
                chips.append((t, f"{TYPES[t]['short']} → {self.sel[t]}" if self.sel[t] else f"{TYPES[t]['short']} · {cnt}"))
        elif st == "tools" and not self.busy():
            for g in ("image", "video", "audio", "pdf"):
                cnt = sum(1 for f in self.files if tool_group(f) == g)
                if cnt: chips.append((g, f"{TOOL_NAMES[g]} · {cnt}"))
        elif self.flash: segs = [(self.flash, "#f4f4f4")]
        elif st == "idle": segs = [(("Брось фото или PDF - сожмём, повернём, склеим" if self.mode == "tools" else f"{len(ALL_EXT)} форматов: фото, видео, аудио, файлы"), COL["muted"])]
        elif st == "chosen": segs = [("Выбери формат · ", COL["muted"]), (save_text(), "#f4f4f4")]
        elif st == "running":
            cur = self.jview.get("cur") or (self.files[0]["name"] if self.files else "")
            segs = [(f"{cur} → {Path(cur).with_suffix('.' + self.target.lower()).name}" if self.target else cur, "#dddddd")]
        elif st == "done": segs = [(self.msg, COL["muted"])]
        elif st == "error": segs = [(self.msg, COL["danger"])]
        f = mkfont(8); fm = QFontMetricsF(f); cf = mkfont(7, True, 0.8); cfm = QFontMetricsF(cf)
        if chips: widths = [cfm.horizontalAdvance(t.upper()) + 18 for _, t in chips]; left = sum(widths) + 4 * (len(chips) - 1)
        else:
            full = "".join(t for t, _ in segs); left = min(290.0, fm.horizontalAdvance(full) + 4)
        total = 16 + left + (8 if btns else 0) + len(btns) * 28 + 4
        total = min(total, W - 12.0); x0 = (W - total) / 2
        self.pill(p, QRectF(x0, BAR_Y, total, BAR_H), "#0c0c0c", "#2e2e2e")
        x = x0 + 16; cy = BAR_Y + BAR_H / 2
        if chips:
            for (t, text), w in zip(chips, widths):
                r = QRectF(x, cy - 11, w, 22); on = t == self.tab; ok = bool(self.sel.get(t))
                self.pill(p, r, "#f4f4f4" if on else "#0c0c0c", "#f4f4f4" if on else ("#8a8a8a" if ok else "#303030"))
                self.dtext(p, text.upper(), r.center().x(), r.center().y(), cf, "#080808" if on else ("#ffffff" if ok else "#9a9a9a"))
                self.hits.append((r, ("chip", t))); x += w + 4
        else:
            avail = left;
            for text, colr in segs:
                text_w = fm.horizontalAdvance(text) + 4
                if x - (x0 + 16) >= avail: break
                w = min(text_w, avail - (x - x0 - 16)); p.setFont(f); p.setPen(QColor(colr))
                p.drawText(QPointF(x, cy + (fm.ascent() - fm.descent()) / 2), fm.elidedText(text, Qt.TextElideMode.ElideRight, w)); x += text_w
        bx = x0 + total - 4 - len(btns) * 28
        for name in btns:
            r = QRectF(bx, cy - 13, 26, 26); hov = self.hover == ("btn", name)
            self.pill(p, r, "#222222" if hov else "#0c0c0c", "#2e2e2e")
            icon = {"cancel": "x", "reveal": "folder", "open": "open", "settings": "settings", "close": "x"}[name]
            draw_icon(p, icon, r.center().x(), r.center().y(), 13, "#ffffff" if hov else "#8a8a8a", 1.9)
            self.hits.append((r, ("btn", name))); bx += 28
        p.restore()

    def mouseMoveEvent(self, e):
        pos = e.position(); h = self.hit_at(pos.x(), pos.y())
        if self.press is None:
            if h != self.hover: self.hover = h; self.update()
            self.setCursor(Qt.CursorShape.PointingHandCursor if (h and not (h[0] == "center" and self.busy())) else Qt.CursorShape.ArrowCursor)

    def leaveEvent(self, e):
        self.hover = None; self.update()

    def mousePressEvent(self, e):
        if e.button() != Qt.MouseButton.LeftButton: return
        pos = e.position(); h = self.hit_at(pos.x(), pos.y())
        if h is None:
            wh = self.windowHandle()
            if wh is not None: wh.startSystemMove()
            return
        self.press = h

    def mouseReleaseEvent(self, e):
        if e.button() != Qt.MouseButton.LeftButton or self.press is None: return
        pos = e.position(); h = self.hit_at(pos.x(), pos.y()); pr = self.press; self.press = None
        if h == pr: self.activate(*pr)

    def activate(self, kind, arg):
        if kind == "tab": self.set_mode(arg)
        elif kind == "btn":
            if arg == "settings": self.ctrl.open_settings()
            elif arg == "close": self.hide_app()
            elif arg == "cancel": self.cancel()
            elif arg == "reveal" and self.last_out: reveal(self.last_out)
            elif arg == "open" and self.last_out: open_path(self.last_out) if Path(self.last_out).is_file() else reveal(self.last_out)
        elif kind == "go": self.start_batch()
        elif kind == "chip": self.tab = arg; self.update()
        elif kind == "fmt":
            if self.mode == "single" and self.state == "chosen": self.run_single(arg)
            elif self.mode == "tools" and self.state == "tools": self.run_tool(arg)
            elif self.mode == "batch" and self.state == "batch" and self.tab:
                self.sel[self.tab] = TYPES[self.tab]["to"][arg]
                nx = next((t for t in self.sel if not self.sel[t]), None)
                if nx: self.tab = nx
                self.update()
        elif kind == "center":
            if self.busy(): return
            if self.state in ("done", "error"): self.reset()
            else: self.choose_dialog()

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls() and not self.busy(): e.acceptProposedAction(); self.over = True; self.update()

    def dragLeaveEvent(self, e):
        self.over = False; self.update()

    def dropEvent(self, e):
        self.over = False; fs = []
        for u in e.mimeData().urls():
            q = Path(u.toLocalFile())
            if q.is_dir(): fs += [x for x in (stat_path(c) for c in sorted(q.iterdir())) if x and x["type"]]
            else:
                x = stat_path(q)
                if x: fs.append(x)
        e.acceptProposedAction(); self.take_files(fs); self.update()

    def keyPressEvent(self, e):
        k = e.key()
        if k == Qt.Key.Key_Escape: self.hide_app(); return
        if k == Qt.Key.Key_F and (e.modifiers() & Qt.KeyboardModifier.ControlModifier): self._prefix = True; self.prefix_t.start(1800); return
        if self._prefix and k in (Qt.Key.Key_H, Qt.Key.Key_V):
            self._prefix = False; self.set_mode("single" if k == Qt.Key.Key_H else "batch"); return
        if k == Qt.Key.Key_V and (e.modifiers() & Qt.KeyboardModifier.ControlModifier): self.paste(); return
        if k in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if self.state == "chosen" and self.mode == "single":
                lf = self.last_fmt(); fm = self.ring_formats()[0]
                if lf in fm: self.run_single(fm.index(lf)); return
            elif self.state == "batch" and self.sel and all(self.sel.values()): self.start_batch(); return
        super().keyPressEvent(e)

    def hide_app(self):
        if not self.isVisible(): return
        self.animate(-1, "hide" if self.ctrl.tray_ok else "quit")

    def closeEvent(self, e):
        if self.ctrl.tray_ok and not self.ctrl.quitting: e.ignore(); self.hide_app()
        else: e.accept()

class NativeHotkeys:
    def __init__(self, cb, combos=None): self.cb = cb; self.ok = False; self.err = ""; self.tr = ComboTracker(combos)
    def start(self):
        ev = threading.Event()
        def run():
            try:
                import ctypes
                from ctypes import wintypes as wt
                u = ctypes.WinDLL("user32", use_last_error=True); k = ctypes.WinDLL("kernel32", use_last_error=True)
                LRESULT = ctypes.c_ssize_t
                PROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, wt.WPARAM, wt.LPARAM)
                class KBD(ctypes.Structure):
                    _fields_ = [("vkCode", wt.DWORD), ("scanCode", wt.DWORD), ("flags", wt.DWORD), ("time", wt.DWORD), ("extra", ctypes.c_size_t)]
                u.SetWindowsHookExW.argtypes = [ctypes.c_int, PROC, ctypes.c_void_p, wt.DWORD]; u.SetWindowsHookExW.restype = ctypes.c_void_p
                u.CallNextHookEx.argtypes = [ctypes.c_void_p, ctypes.c_int, wt.WPARAM, wt.LPARAM]; u.CallNextHookEx.restype = LRESULT
                u.GetAsyncKeyState.argtypes = [ctypes.c_int]; u.GetAsyncKeyState.restype = ctypes.c_short
                u.GetMessageW.argtypes = [ctypes.POINTER(wt.MSG), ctypes.c_void_p, wt.UINT, wt.UINT]
                k.GetModuleHandleW.argtypes = [wt.LPCWSTR]; k.GetModuleHandleW.restype = ctypes.c_void_p
                self.tr.is_pressed = lambda vk: u.GetAsyncKeyState(vk) & 0x8000
                def proc(n, w, l):
                    try:
                        if n == 0 and w in (0x100, 0x101, 0x104, 0x105):
                            vk = ctypes.cast(l, ctypes.POINTER(KBD)).contents.vkCode
                            act, sup = self.tr.feed(vk, w in (0x100, 0x104))
                            if act: self.cb(act)
                            if sup: return 1
                    except Exception: pass
                    return u.CallNextHookEx(None, n, w, l)
                self._proc = PROC(proc)
                h = u.SetWindowsHookExW(13, self._proc, k.GetModuleHandleW(None), 0)
                if not h: raise OSError("SetWindowsHookEx: код %s" % ctypes.get_last_error())
                self.ok = True; ev.set()
                msg = wt.MSG()
                while u.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                    u.TranslateMessage(ctypes.byref(msg)); u.DispatchMessageW(ctypes.byref(msg))
            except Exception as e:
                self.err = short_err(e); log("Горячие клавиши (системный перехват):", traceback.format_exc()); ev.set()
        threading.Thread(target=run, daemon=True).start(); ev.wait(3)
        return self.ok

class HotkeyButton(QPushButton):
    def __init__(self, combo):
        super().__init__()
        self.combo = combo; self.rec = False; self.held = []; self.seen = []
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus); self.setMinimumWidth(200)
        self.clicked.connect(self.start); self.refresh()

    def refresh(self, text=None):
        self.setText(text or ("Нажми сочетание...  (Esc - отмена)" if self.rec else self.combo.replace("+", "  +  ")))

    def set_combo(self, combo): self.combo = combo; self.stop()

    def start(self):
        self.rec = True; self.held = []; self.seen = []; self.refresh(); self.setFocus(); self.grabKeyboard()

    def stop(self, text=None):
        if self.rec: self.releaseKeyboard()
        self.rec = False; self.refresh(text)

    def focusOutEvent(self, e):
        if self.rec: self.stop()
        super().focusOutEvent(e)

    def focusNextPrevChild(self, nxt):
        return False if self.rec else super().focusNextPrevChild(nxt)

    def event(self, e):
        if self.rec:
            t = e.type()
            if t == QEvent.Type.ShortcutOverride: e.accept(); return True
            if t == QEvent.Type.KeyPress: self.key(e, True); return True
            if t == QEvent.Type.KeyRelease: self.key(e, False); return True
        return super().event(e)

    def key(self, e, down):
        if e.isAutoRepeat(): return
        if down and e.key() == Qt.Key.Key_Escape and not self.seen: self.stop(); return
        tok = vk_token(e.nativeVirtualKey())
        if down:
            if tok is None: self.refresh("Эта клавиша не поддерживается"); return
            if tok not in self.held: self.held.append(tok)
            if tok not in self.seen: self.seen.append(tok)
            self.refresh("  +  ".join(m for m in MOD_ORDER if m in self.seen) + ("  +  " if any(m in self.seen for m in MOD_ORDER) else "") + "  +  ".join(t for t in self.seen if t not in VK_MODS) + "  ...")
        else:
            if tok in self.held: self.held.remove(tok)
            if not self.held and self.seen:
                combo = normalize_combo(self.seen); err = combo_error(combo); old = self.combo; self.seen = []
                if err: self.stop(err); QTimer.singleShot(1800, self.refresh); return
                self.combo = combo; self.stop()

QSS = """
* { font-family: 'Segoe UI', 'Arial', sans-serif; color: #f4f4f4; font-size: 12px; }
QDialog { background: #0b0b0b; }
QLabel#Eyebrow { color: #929292; font-size: 9px; font-weight: 600; letter-spacing: 1.5px; }
QLabel#Title { font-size: 22px; font-weight: 600; }
QLabel#Muted { color: #929292; font-size: 11px; }
QLabel#Info { color: #b8b8b8; background: #151515; border: 1px solid #292929; border-radius: 10px; padding: 11px; font-size: 11px; }
QPushButton { background: #171717; color: #eaeaea; border: 1px solid #303030; border-radius: 10px; padding: 9px 14px; }
QPushButton:hover { background: #222222; border-color: #4a4a4a; }
QPushButton#Primary { background: #f4f4f4; color: #080808; border: 1px solid #ffffff; font-weight: 700; }
QPushButton#Primary:hover { background: #dcdcdc; }
QCheckBox, QRadioButton { color: #c6c6c6; spacing: 9px; font-size: 12px; }
QCheckBox::indicator { width: 15px; height: 15px; border: 1px solid #4a4a4a; border-radius: 4px; background: #111111; }
QCheckBox::indicator:checked { background: #f4f4f4; border-color: #f4f4f4; image: url(%CHECK%); }
QRadioButton::indicator { width: 15px; height: 15px; border: 1px solid #4a4a4a; border-radius: 8px; background: #111111; }
QRadioButton::indicator:checked { background: #f4f4f4; border-color: #f4f4f4; image: url(%DOT%); }
QCheckBox:disabled { color: #5e5e5e; }
QPushButton#Hotkey { text-align: center; font-weight: 600; }
QMenu { background: #151515; border: 1px solid #303030; padding: 4px; }
QMenu::item { padding: 7px 22px; border-radius: 6px; }
QMenu::item:selected { background: #2a2a2a; }
QMenu::separator { height: 1px; background: #2a2a2a; margin: 4px 6px; }
QLabel#Dim { color: #5e5e5e; font-size: 11px; }
QScrollArea, QScrollArea > QWidget > QWidget { background: #0b0b0b; border: none; }
QSpinBox, QLineEdit, QComboBox { background: #141414; border: 1px solid #303030; border-radius: 8px; padding: 6px 9px; color: #f4f4f4; selection-background-color: #f4f4f4; selection-color: #080808; }
QSpinBox:focus, QLineEdit:focus, QComboBox:focus { border: 1px solid #8a8a8a; }
QSpinBox::up-button, QSpinBox::down-button { width: 0; border: none; }
QComboBox::drop-down { border: none; width: 24px; }
QComboBox::down-arrow { image: url(%ARROW%); width: 10px; height: 10px; }
QComboBox QAbstractItemView { background: #141414; border: 1px solid #303030; selection-background-color: #2a2a2a; selection-color: #ffffff; color: #f4f4f4; outline: 0; }
QSlider::groove:horizontal { height: 4px; background: #2a2a2a; border-radius: 2px; }
QSlider::sub-page:horizontal { background: #f4f4f4; border-radius: 2px; }
QSlider::handle:horizontal { background: #f4f4f4; width: 16px; height: 16px; margin: -6px 0; border-radius: 8px; }
QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: #3a3a3a; border-radius: 4px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #5a5a5a; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QToolTip { background: #151515; color: #f4f4f4; border: 1px solid #303030; padding: 4px; }
QMessageBox { background: #0b0b0b; }
QDialog#Panel { background: transparent; }
QPushButton#XBtn { background: #0c0c0c; border: 1px solid #2e2e2e; border-radius: 15px; padding: 0; font-size: 16px; color: #9a9a9a; }
QPushButton#XBtn:hover { background: #222222; color: #ffffff; }
QPushButton#TabBtn { background: #0c0c0c; border: 1px solid #2c2c2c; border-radius: 15px; padding: 6px 12px; font-size: 11px; color: #9a9a9a; }
QPushButton#TabBtn:hover { color: #ffffff; }
QPushButton#TabBtn:checked { background: #f4f4f4; color: #080808; border-color: #f4f4f4; font-weight: 600; }
QLabel#Sub { color: #7a7a7a; font-size: 11px; }
"""

class Panel(QDialog):
    R = 26
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Panel"); self.setWindowIcon(qicon()); self.setModal(True)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self._x = QPushButton("×", self); self._x.setObjectName("XBtn"); self._x.setFixedSize(30, 30); self._x.clicked.connect(self.reject)

    def resizeEvent(self, e):
        self._x.move(self.width() - 48, 18); self._x.raise_(); super().resizeEvent(e)

    def paintEvent(self, e):
        p = QPainter(self); p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        g = QLinearGradient(0, 0, 0, self.height()); g.setColorAt(0, QColor("#171717")); g.setColorAt(1, QColor("#0a0a0a"))
        p.setPen(QPen(QColor("#2a2a2a"), 1.2)); p.setBrush(QBrush(g)); p.drawRoundedRect(QRectF(1, 1, self.width() - 2, self.height() - 2), self.R, self.R); p.end()

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            ch = self.childAt(e.position().toPoint())
            if ch is None or isinstance(ch, QLabel):
                wh = self.windowHandle()
                if wh is not None: wh.startSystemMove()

def settings_qss():
    try:
        from PIL import Image, ImageDraw
        DATA.mkdir(parents=True, exist_ok=True)
        c = Image.new("RGBA", (60, 60), (0, 0, 0, 0)); ImageDraw.Draw(c).line([(14, 31), (25, 42), (46, 18)], fill=(8, 8, 8, 255), width=7, joint="curve")
        c.resize((15, 15), Image.LANCZOS).save(DATA / "chk.png")
        d = Image.new("RGBA", (60, 60), (0, 0, 0, 0)); ImageDraw.Draw(d).ellipse((18, 18, 42, 42), fill=(8, 8, 8, 255))
        d.resize((15, 15), Image.LANCZOS).save(DATA / "dot.png")
        a = Image.new("RGBA", (60, 60), (0, 0, 0, 0)); ImageDraw.Draw(a).line([(12, 22), (30, 40), (48, 22)], fill=(230, 230, 230, 255), width=7, joint="curve")
        a.resize((10, 10), Image.LANCZOS).save(DATA / "arr.png")
        return QSS.replace("%CHECK%", (DATA / "chk.png").as_posix()).replace("%DOT%", (DATA / "dot.png").as_posix()).replace("%ARROW%", (DATA / "arr.png").as_posix())
    except Exception: return QSS.replace("image: url(%CHECK%);", "").replace("image: url(%DOT%);", "").replace("image: url(%ARROW%);", "")

class ToolDialog(Panel):
    TITLES = {"compress": "Сжать фото", "resize": "Изменить размер", "rotate": "Повернуть фото", "crop": "Обрезать по формату", "watermark": "Подпись на фото",
              "pdf_split": "Разбить PDF", "pdf_rotate": "Повернуть PDF", "pdf_compress": "Сжать PDF", "vid_compress": "Сжать видео", "trim": "Обрезать по времени"}
    POS = ["br", "bl", "tr", "tl", "center", "tile"]

    def __init__(self, op, parent=None):
        super().__init__(parent)
        self.op = op; self.setWindowTitle(self.TITLES.get(op, "Параметры")); self.setMinimumWidth(440)
        L = QVBoxLayout(self); L.setContentsMargins(32, 30, 32, 26); L.setSpacing(10)
        t = QLabel(self.TITLES.get(op, "Параметры")); t.setObjectName("Title"); L.addWidget(t)
        def radio(text, on=False): r = QRadioButton(text); r.setChecked(on); L.addWidget(r); return r
        def spin(lo, hi, val, suf=""):
            s = QSpinBox(); s.setRange(lo, hi); s.setValue(val); s.setSuffix(suf); s.setMinimumWidth(110); return s
        def line(*ws):
            h = QHBoxLayout(); [h.addWidget(w) for w in ws]; h.addStretch(1); L.addLayout(h)
        def slider(lo, hi, val):
            s = QSlider(Qt.Orientation.Horizontal); s.setRange(lo, hi); s.setValue(val); return s
        if op == "compress":
            self.r1 = radio("По качеству", True); self.sl = slider(10, 95, 75); self.lb = QLabel("75"); self.sl.valueChanged.connect(lambda v: self.lb.setText(str(v))); line(self.sl, self.lb)
            self.r2 = radio("До заданного размера"); self.kb = spin(10, 50000, 500, " КБ"); line(self.kb)
        elif op == "resize":
            self.r1 = radio("В процентах", True); self.pc = spin(5, 400, 50, " %"); line(self.pc)
            self.r2 = radio("Вписать в размер (0 - не ограничивать)"); self.w = spin(0, 20000, 1280, " px"); self.h = spin(0, 20000, 0, " px"); line(QLabel("ширина"), self.w, QLabel("высота"), self.h)
        elif op == "rotate":
            self.opts = [(radio("На 90° вправо", True), "cw90"), (radio("На 90° влево"), "ccw90"), (radio("На 180°"), "180"), (radio("Отразить по горизонтали"), "fliph"), (radio("Отразить по вертикали"), "flipv")]
        elif op == "crop":
            self.opts = [(radio("1:1 - квадрат (аватар, соцсети)", True), "1:1"), (radio("4:3 - классический снимок"), "4:3"), (radio("3:2 - кадр с фотоаппарата"), "3:2"), (radio("16:9 - экран и YouTube"), "16:9"), (radio("9:16 - вертикальное видео, сторис"), "9:16")]
        elif op == "watermark":
            L.addWidget(QLabel("Текст подписи")); self.ed = QLineEdit(); self.ed.setText("© " + time.strftime("%Y")); L.addWidget(self.ed)
            L.addWidget(QLabel("Где")); self.cb = QComboBox(); self.cb.addItems(["Справа внизу", "Слева внизу", "Справа вверху", "Слева вверху", "По центру", "Мозаикой по всему фото"]); L.addWidget(self.cb)
            L.addWidget(QLabel("Непрозрачность")); self.op_ = slider(10, 100, 60); L.addWidget(self.op_)
            L.addWidget(QLabel("Размер текста")); self.sz = slider(2, 15, 5); L.addWidget(self.sz)
        elif op == "pdf_split":
            self.r1 = radio("Каждая страница - отдельный файл", True); self.r2 = radio("По диапазонам, через запятую")
            self.ed = QLineEdit(); self.ed.setPlaceholderText("например 1-3,5,7-9"); L.addWidget(self.ed)
        elif op == "pdf_rotate":
            self.opts = [(radio("На 90° вправо", True), 90), (radio("На 90° влево"), 270), (radio("На 180°"), 180)]
            L.addWidget(QLabel("Страницы (пусто - все):")); self.ed = QLineEdit(); self.ed.setPlaceholderText("например 1-3,5"); L.addWidget(self.ed)
        elif op == "pdf_compress":
            self.opts = [(radio("Лёгкое - без потери качества"), 1), (radio("Среднее - заметно меньше", True), 2), (radio("Сильное - самый маленький файл"), 3)]
        elif op == "vid_compress":
            self.opts = [(radio("Почти без потери качества (для хранения)"), "hq"), (radio("Для телефона и почты - до 720p", True), "mid"), (radio("Для мессенджеров - до 480p, самый маленький"), "low")]
            self.r2 = radio("До заданного размера"); self.mb = spin(1, 20000, 50, " МБ"); line(self.mb)
        elif op == "trim":
            L.addWidget(QLabel("Начало (например 0:30 или 90)")); self.s0 = QLineEdit(); self.s0.setPlaceholderText("00:00"); L.addWidget(self.s0)
            L.addWidget(QLabel("Конец (пусто - до конца)")); self.s1 = QLineEdit(); self.s1.setPlaceholderText("до конца"); L.addWidget(self.s1)
            self.r1 = radio("Точно - с перекодированием (рекомендуется)", True); self.r2 = radio("Быстро - без перекодирования, начало сдвинется до ключевого кадра")
        row = QHBoxLayout(); row.addStretch(1); c = QPushButton("Отмена"); ok = QPushButton("Применить"); ok.setObjectName("Primary")
        c.clicked.connect(self.reject); ok.clicked.connect(self.accept); row.addWidget(c); row.addWidget(ok); L.addLayout(row)

    def params(self):
        o = self.op; pick = lambda: next(v for r, v in self.opts if r.isChecked())
        if o == "compress": return dict(quality=self.sl.value()) if self.r1.isChecked() else dict(target_kb=self.kb.value())
        if o == "resize": return dict(mode="percent", value=self.pc.value()) if self.r1.isChecked() else dict(mode="box", w=self.w.value(), h=self.h.value())
        if o in ("rotate", "crop"): return dict(angle=pick()) if o == "rotate" else dict(ratio=pick())
        if o == "watermark": return dict(text=self.ed.text(), pos=self.POS[self.cb.currentIndex()], opacity=self.op_.value(), size=self.sz.value())
        if o == "pdf_split": return dict(mode="each") if self.r1.isChecked() else dict(mode="ranges", ranges=self.ed.text())
        if o == "pdf_rotate": return dict(angle=pick(), pages=self.ed.text())
        if o == "pdf_compress": return dict(level=pick())
        if o == "vid_compress": return dict(target_mb=self.mb.value()) if self.r2.isChecked() else dict(preset=pick())
        if o == "trim": return dict(start=self.s0.text(), end=self.s1.text(), mode="exact" if self.r1.isChecked() else "fast")
        return {}

def ask_tool_params(op, parent):
    if op not in ToolDialog.TITLES: return {}
    d = ToolDialog(op, parent)
    return d.params() if d.exec() else None

class HistoryDialog(Panel):
    def __init__(self, ctrl, parent=None):
        super().__init__(parent)
        self.setWindowTitle("История · " + APP); self.setFixedSize(600, 560)
        outer = QVBoxLayout(self); outer.setContentsMargins(32, 30, 32, 26); outer.setSpacing(10)
        h = QLabel("История"); h.setObjectName("Title"); outer.addWidget(h)
        inner = QWidget(); L = QVBoxLayout(inner); L.setContentsMargins(0, 0, 8, 0); L.setSpacing(10)
        sc = QScrollArea(); sc.setWidgetResizable(True); sc.setFrameShape(QFrame.Shape.NoFrame); sc.setWidget(inner); outer.addWidget(sc, 1)
        sc.setStyleSheet("QScrollArea, QScrollArea > QWidget > QWidget { background: transparent; }"); sc.viewport().setAutoFillBackground(False)
        items = history_load()
        if not items:
            m = QLabel("Здесь появятся последние конвертации и операции."); m.setObjectName("Muted"); L.addWidget(m)
        for e in items:
            row = QHBoxLayout(); dst = Path(e["dst"]); ex = dst.exists()
            t = QLabel(f"{e['src']}  →  {dst.name}\n{e['what']} · {fmt_size(e['si'])} → {fmt_size(e['so'])} · {time.strftime('%d.%m %H:%M', time.localtime(e['t']))}")
            t.setObjectName("Muted" if ex else "Dim"); t.setWordWrap(True)
            b1 = QPushButton("Открыть"); b2 = QPushButton("Папка"); b1.setEnabled(ex); b2.setEnabled(ex)
            b1.clicked.connect(lambda _=False, q=dst: open_path(q)); b2.clicked.connect(lambda _=False, q=dst: reveal(q))
            row.addWidget(t, 1); row.addWidget(b1); row.addWidget(b2); L.addLayout(row)
        L.addStretch(1)
        foot = QHBoxLayout(); foot.addStretch(1)
        cl = QPushButton("Очистить"); ok = QPushButton("Закрыть"); ok.setObjectName("Primary")
        cl.clicked.connect(lambda: (history_clear(), self.accept())); ok.clicked.connect(self.accept); foot.addWidget(cl); foot.addWidget(ok); outer.addLayout(foot)

class SettingsDialog(Panel):
    PAGES = ("Сохранение", "Качество", "Клавиши", "Папка", "Система")

    def __init__(self, ctrl, parent=None):
        super().__init__(parent)
        self.ctrl = ctrl; self.setWindowTitle("Настройки · " + APP); self.setFixedSize(520, 600)
        def lab(text, name):
            l = QLabel(text); l.setObjectName(name); l.setWordWrap(True); return l
        L = QVBoxLayout(self); L.setContentsMargins(32, 30, 32, 26); L.setSpacing(12)
        L.addWidget(lab("Настройки", "Title"))
        tabs = QHBoxLayout(); tabs.setSpacing(6); self.tb = []; self.stack = QStackedWidget()
        for i, name in enumerate(self.PAGES):
            b = QPushButton(name); b.setObjectName("TabBtn"); b.setCheckable(True); b.clicked.connect(lambda _=False, i=i: self.go(i)); tabs.addWidget(b); self.tb.append(b)
        tabs.addStretch(1); L.addLayout(tabs); L.addWidget(self.stack, 1)
        def page():
            w = QWidget(); v = QVBoxLayout(w); v.setContentsMargins(0, 8, 0, 0); v.setSpacing(9); self.stack.addWidget(w); return v
        def opt(v, widget, desc=""):
            v.addWidget(widget)
            if desc: d = lab(desc, "Sub"); d.setContentsMargins(26, 0, 0, 0); v.addWidget(d)

        v = page(); m = S.get("mode", "near")
        self.near = QRadioButton("Рядом с файлом"); self.fold = QRadioButton("В отдельную папку"); self.repl = QRadioButton("Вместо оригинала")
        self.near.setChecked(m not in ("folder", "replace")); self.fold.setChecked(m == "folder"); self.repl.setChecked(m == "replace")
        self.path = S.get("outdir_path") or str(default_outdir()); self.plab = lab(self.path, "Sub"); self.plab.setContentsMargins(26, 0, 0, 0); self.pbtn = QPushButton("Выбрать папку...")
        def pick():
            d = QFileDialog.getExistingDirectory(self, "Папка для результатов", self.path)
            if d: self.path = d; self.plab.setText(d)
        self.pbtn.clicked.connect(pick)
        self.trash = QCheckBox("Исходный файл отправлять в корзину"); self.trash.setChecked(bool(S.get("trash", True)))
        opt(v, self.near, "Новый файл появится в той же папке, оригинал останется на месте")
        opt(v, self.fold, "Все результаты будут собираться в одной папке"); v.addWidget(self.plab); v.addWidget(self.pbtn)
        opt(v, self.repl, "Результат встанет на место исходного файла"); v.addWidget(self.trash); v.addStretch(1)
        self.fold.toggled.connect(lambda on: (self.pbtn.setEnabled(on), self.plab.setEnabled(on))); self.pbtn.setEnabled(self.fold.isChecked()); self.plab.setEnabled(self.fold.isChecked())
        self.repl.toggled.connect(lambda on: self.trash.setEnabled(on)); self.trash.setEnabled(self.repl.isChecked())

        v = page()
        self.q = QSlider(Qt.Orientation.Horizontal); self.q.setRange(40, 100); self.q.setValue(int(S.get("quality", 90)))
        self.ql = QLabel(); self.q.valueChanged.connect(lambda x: self.ql.setText(f"Качество фото: {x}")); self.ql.setText(f"Качество фото: {self.q.value()}")
        v.addWidget(self.ql); v.addWidget(self.q); v.addWidget(lab("Выше - лучше картинка, но файл больше. Для JPG, WebP и AVIF.", "Sub"))
        self.dpi = QSpinBox(); self.dpi.setRange(50, 600); self.dpi.setValue(int(S.get("dpi", 150))); self.dpi.setSuffix(" DPI"); self.dpi.setMinimumWidth(120)
        dr = QHBoxLayout(); dr.addWidget(QLabel("Чёткость PDF и страниц:")); dr.addWidget(self.dpi); dr.addStretch(1); v.addSpacing(8); v.addLayout(dr)
        v.addWidget(lab("150 хватает для экрана, 300 - для печати.", "Sub"))
        self.merge = QCheckBox("Фото в PDF склеивать в один файл"); self.merge.setChecked(bool(S.get("merge_pdf", True))); v.addSpacing(8)
        opt(v, self.merge, "Если выбрано несколько фото и формат PDF, получится один документ"); v.addStretch(1)

        v = page(); self.hk = {}
        for key, name in (("hk_single", "Открыть: один файл"), ("hk_batch", "Открыть: несколько файлов")):
            v.addWidget(QLabel(name)); row = QHBoxLayout(); b = HotkeyButton(S.get(key) or DEFAULTS[key]); b.setObjectName("Hotkey")
            r = QPushButton("Сброс"); r.clicked.connect(lambda _=False, b=b, k=key: b.set_combo(DEFAULTS[k])); row.addWidget(b, 1); row.addWidget(r); v.addLayout(row); self.hk[key] = b
        v.addWidget(lab("Нажми на кнопку и набери сочетание: Ctrl, Alt или Shift плюс клавиши. Повторное нажатие закрывает окно." + ("" if ctrl.hk_on else " Сейчас горячие клавиши не работают - причина в журнале ошибок (меню значка)."), "Sub")); v.addStretch(1)

        v = page(); self.wc = {}
        self.won = QCheckBox("Следить за папкой"); self.won.setChecked(bool(S.get("watch_on")))
        opt(v, self.won, "Новые файлы в ней будут конвертироваться сами. Старые не трогаются.")
        self.wpath = S.get("watch_dir") or str(Path.home() / "Downloads"); self.wlab = lab(self.wpath, "Sub"); wb = QPushButton("Выбрать папку...")
        def wpick():
            d = QFileDialog.getExistingDirectory(self, "Папка для слежения", self.wpath)
            if d: self.wpath = d; self.wlab.setText(d)
        wb.clicked.connect(wpick); v.addWidget(self.wlab); v.addWidget(wb)
        for key, name, opts in (("watch_img", "Фото сохранять как", ["JPG", "PNG", "WEBP"]), ("watch_vid", "Видео сохранять как", ["MP4", "MKV"]), ("watch_aud", "Аудио сохранять как", ["MP3", "M4A", "OGG", "FLAC"]), ("watch_doc", "Документы сохранять как", ["PDF", "DOCX", "TXT"])):
            cb = QComboBox(); cb.addItems(["не трогать"] + opts); cur = S.get(key, ""); cb.setCurrentIndex(opts.index(cur) + 1 if cur in opts else 0)
            r = QHBoxLayout(); r.addWidget(QLabel(name)); r.addWidget(cb, 1); v.addLayout(r); self.wc[key] = (cb, opts)
        v.addStretch(1)

        v = page()
        self.auto = QCheckBox("Запускать вместе с Windows"); self.auto.setChecked(bool(S.get("autostart")))
        self.ctx = QCheckBox("Пункт «Конвертировать в...» в меню правой кнопки"); self.ctx.setChecked(bool(S.get("ctx")))
        self.notif = QCheckBox("Сообщать о завершении"); self.notif.setChecked(bool(S.get("notify", True)))
        self.office = QCheckBox("Использовать Word, PowerPoint и LibreOffice"); self.office.setChecked(bool(S.get("office", True)))
        for w in (self.auto, self.ctx): w.setEnabled(IS_WIN)
        opt(v, self.auto, "Программа будет тихо ждать в трее"); opt(v, self.ctx, "Быстрая конвертация прямо из проводника")
        opt(v, self.notif, "Уведомление рядом с часами"); opt(v, self.office, "Нужны для старых форматов вроде .doc и .ppt, если они установлены")
        self.info = lab(self.engine_text(), "Sub"); v.addSpacing(6); v.addWidget(self.info)
        hr = QHBoxLayout(); hb = QPushButton("История..."); hb.clicked.connect(lambda: HistoryDialog(self.ctrl, self).exec())
        rb = QPushButton("Проверить программы"); rb.clicked.connect(lambda: self.info.setText(self.engine_text())); hr.addWidget(hb); hr.addWidget(rb); v.addLayout(hr); v.addStretch(1)

        row = QHBoxLayout(); row.addStretch(1); cancel = QPushButton("Отмена"); ok = QPushButton("Сохранить"); ok.setObjectName("Primary")
        cancel.clicked.connect(self.reject); ok.clicked.connect(self.save); row.addWidget(cancel); row.addWidget(ok); L.addLayout(row)
        self.go(0)

    def go(self, i):
        self.stack.setCurrentIndex(i)
        for k, b in enumerate(self.tb): b.setChecked(k == i)

    def engine_text(self):
        e = engines(); yn = lambda v: "есть" if v else "нет"
        return f"Форматов: {len(ALL_EXT)}. FFmpeg: {yn(e['ffmpeg'])} · LibreOffice: {yn(e['libre'])} · Word: {yn(e['word'])} · PowerPoint: {yn(e['ppt'])} · HEIC: {yn(e['heif'])}"

    def save(self):
        a, b = self.hk["hk_single"].combo, self.hk["hk_batch"].combo
        if combos_conflict(a, b): QMessageBox.warning(self, APP, "Сочетания для «Один файл» и «Несколько файлов» не должны совпадать или входить друг в друга."); return
        want = dict(**{k: (o[cb.currentIndex() - 1] if cb.currentIndex() else "") for k, (cb, o) in self.wc.items()}, watch_on=self.won.isChecked(), watch_dir=self.wpath, hk_single=a, hk_batch=b, mode="replace" if self.repl.isChecked() else "folder" if self.fold.isChecked() else "near", outdir_path=self.path,
                    quality=self.q.value(), dpi=self.dpi.value(), merge_pdf=self.merge.isChecked(), trash=self.trash.isChecked(), notify=self.notif.isChecked(), office=self.office.isChecked())
        if IS_WIN: want.update(autostart=self.auto.isChecked(), ctx=self.ctx.isChecked())
        for k, v in want.items():
            if S.get(k) == v: continue
            err = apply_setting(k, v)
            if err: QMessageBox.warning(self, APP, err)
        save_settings(); self.ctrl.reload_hotkeys(); self.ctrl.reload_watch(); self.accept()

class Controller(QObject):
    sig_open = Signal(str)

    def __init__(self, app, args):
        super().__init__()
        self.app = app; self.args = args; self.quitting = False; self.hk_on = False; self.tray_ok = False
        self.win = RoundWindow(self); self.pend = []
        self.pend_t = QTimer(self); self.pend_t.setSingleShot(True); self.pend_t.setInterval(700); self.pend_t.timeout.connect(self.flush)
        self.sig_open.connect(self._on_open)
        self.watch = Watcher(); self.wjobs = []
        self.wt = QTimer(self); self.wt.setInterval(3000); self.wt.timeout.connect(self._watch_tick); self.wt.start()
        self.setup_tray(); self.setup_server(); self.setup_hotkeys()

    def _watch_tick(self):
        try:
            self.wjobs += self.watch.tick()
            for j in list(self.wjobs):
                if j.state == "run": continue
                self.wjobs.remove(j); ok = [r for r in j.results if r.get("ok")]
                self.notify("Автоконвертация: " + Path(ok[-1]["dst"]).name if ok else "Автоконвертация не удалась: " + (j.results[0].get("error", "") if j.results else ""))
        except Exception: log("слежение за папкой:", traceback.format_exc())

    def reload_watch(self): self.watch.armed = None

    def open_history(self):
        self.win.popup(at_cursor=False); HistoryDialog(self, self.win).exec()

    def paste_open(self):
        self.win.open_request("single"); self.win.paste()

    def _on_open(self, mode):
        self.toggle(mode)

    def toggle(self, mode):
        w = self.win
        if w.shown_ok():
            if w.mode == mode: w.hide_app()
            elif not w.busy(): w.set_mode(mode)
        else: w.open_request(mode)

    def setup_tray(self):
        if not QSystemTrayIcon.isSystemTrayAvailable(): return
        self.tray = QSystemTrayIcon(qicon(), self); self.tray.setToolTip(APP)
        m = QMenu()
        self.a1 = m.addAction("Один файл"); self.a2 = m.addAction("Несколько файлов"); a1, a2 = self.a1, self.a2; m.addSeparator()
        a6 = m.addAction("Из буфера обмена  (Ctrl+V)"); a7 = m.addAction("История"); m.addSeparator(); a3 = m.addAction("Настройки"); a5 = m.addAction("Журнал ошибок"); m.addSeparator(); a4 = m.addAction("Выход")
        a1.triggered.connect(lambda: self.win.open_request("single")); a2.triggered.connect(lambda: self.win.open_request("batch"))
        a3.triggered.connect(self.open_settings); a4.triggered.connect(self.quit); a5.triggered.connect(self.open_log); a6.triggered.connect(self.paste_open); a7.triggered.connect(self.open_history)
        self.tray.setContextMenu(m); self._menu = m; self.refresh_labels()
        self.tray.activated.connect(lambda r: self.toggle(self.win.mode) if r in (QSystemTrayIcon.ActivationReason.Trigger, QSystemTrayIcon.ActivationReason.DoubleClick) else None)
        self.tray.show(); self.tray_ok = True

    def refresh_labels(self):
        try:
            c = current_combos(); self.a1.setText(f"Один файл   {c['single']}"); self.a2.setText(f"Несколько файлов   {c['batch']}")
        except Exception: pass

    def reload_hotkeys(self):
        c = current_combos(); self.refresh_labels(); log("Сочетания клавиш:", c)
        try:
            if getattr(self, "_nh", None): self._nh.tr.set_combos(c)
        except Exception: log("Не удалось обновить сочетания:", traceback.format_exc())

    def notify(self, text):
        if S.get("notify", True) and self.tray_ok and not self.win.isActiveWindow():
            try: self.tray.showMessage(APP, text, QSystemTrayIcon.MessageIcon.Information, 4000)
            except Exception: pass

    def open_settings(self):
        self.win.popup(at_cursor=False); self.win.dialog_open = True
        try: SettingsDialog(self, self.win).exec()
        finally: self.win.dialog_open = False; self.win.update()

    def quit(self):
        self.quitting = True
        try:
            if self.win.job: self.win.job.kill()
        except Exception: pass
        self.app.quit()

    def setup_hotkeys(self):
        self.hk_method = ""
        if self.args.get("data"): return
        if IS_WIN:
            nh = NativeHotkeys(lambda m: self.sig_open.emit(m), current_combos())
            if nh.start(): self.hk_on = True; self.hk_method = "Windows"; self._nh = nh; log("Горячие клавиши: системный перехват, OK"); return
            log("Системный перехват клавиш не запустился:", nh.err)
        if not self.hk_on: log("Горячие клавиши НЕ активны")

    def open_log(self):
        try:
            lg = DATA / "convi.log"
            if not lg.exists(): lg.write_text("", "utf-8")
            if IS_WIN: os.startfile(str(lg))
            else: subprocess.Popen(["xdg-open", str(lg)])
        except Exception: pass

    def setup_server(self):
        self.server = QLocalServer(self); name = server_name(self.args)
        QLocalServer.removeServer(name); self.server.listen(name); self.server.newConnection.connect(self._conn)

    def _conn(self):
        while self.server.hasPendingConnections():
            s = self.server.nextPendingConnection(); s.readyRead.connect(lambda s=s: self._read(s))

    def _read(self, s):
        try: data = bytes(s.readAll()).decode("utf-8", "replace")
        except Exception: return
        for line in data.splitlines():
            try: d = json.loads(line)
            except Exception: continue
            self.collect(d.get("action", "files"), d.get("files") or [], d.get("to"))

    def collect(self, action, files, to=None):
        self.pend += [(action, f, to) for f in files] or [(action, None, to)]
        self.pend_t.start()

    def flush(self):
        items = self.pend; self.pend = []
        if not items: return
        action = items[0][0]; to = items[0][2]; files = [f for _, f, _ in items if f]
        st = [x for x in (stat_path(f) for f in files) if x]
        if action == "convert" and len(st) == 1: self.win.open_request("single", st, to)
        elif action == "convert" and st:
            pre = {t: to.upper() for t, v in TYPES.items() if to and to.upper() in v["to"] and any(ftype(f) == t for f in st)}
            self.win.open_request("batch", st, None, pre)
        elif action == "batch": self.win.open_request("batch", st)
        elif action == "single" or not st: self.win.open_request("single", st)
        else: self.win.open_request("single" if len(st) == 1 else "batch", st)

def server_name(args):
    d = args.get("data")
    return "Convi-v" + VERSION.replace(".", "_") + "-" + (re.sub(r"\W", "_", str(d)) if d else "main")

def send_to_running(args):
    s = QLocalSocket(); s.connectToServer(server_name(args))
    if not s.waitForConnected(400): return False
    files = [str(Path(f).resolve()) for f in args["files"]]
    act = "convert" if args["convert"] else "batch" if args["batch"] else "single" if args["single"] else "files"
    s.write((json.dumps(dict(action=act, files=files, to=args["to"])) + "\n").encode("utf-8"))
    s.waitForBytesWritten(1500); s.disconnectFromServer(); return True

def parse_args(argv):
    a = dict(tray=False, setup=False, convert=False, batch=False, single=False, to=None, files=[], data=None, make_icon=None)
    i = 0
    while i < len(argv):
        x = argv[i]
        if x in ("--tray", "--setup", "--convert", "--batch", "--single"): a[x[2:]] = True
        elif x == "--minimized": a["tray"] = True
        elif x == "--to" and i + 1 < len(argv): a["to"] = argv[i + 1]; i += 1
        elif x == "--data" and i + 1 < len(argv): a["data"] = argv[i + 1]; i += 1
        elif x == "--make-icon" and i + 1 < len(argv): a["make_icon"] = argv[i + 1]; i += 1
        elif not x.startswith("--"): a["files"].append(x)
        i += 1
    return a

def main():
    global DATA
    args = parse_args(sys.argv[1:])
    if args["setup"]:
        print("Convi: установка библиотек (один раз)...", flush=True)
        ok = bootstrap(True); print("Готово." if ok else "Есть ошибки установки - проверь интернет и запусти снова.")
        sys.exit(0 if ok else 1)
    if args["make_icon"]:
        bootstrap(False); make_ico(args["make_icon"]); print("Иконка создана:", args["make_icon"]); return
    if not HAVE_QT:
        if os.environ.get("CONVI_REEXEC"): print("Не удалось подключить PySide6. Запусти: python convi.py --setup"); sys.exit(1)
        bootstrap(False); os.environ["CONVI_REEXEC"] = "1"; os.execv(sys.executable, [sys.executable] + sys.argv)
    if args["data"]: DATA = Path(args["data"])
    install_error_hooks(); log(f"=== запуск v{VERSION}, frozen={FROZEN}, Python {sys.version.split()[0]}")
    cleanup_legacy(); load_settings()
    if IS_WIN:
        try:
            import ctypes; ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("CONVI.Convert")
        except Exception: pass
    app = QApplication(sys.argv[:1]); app.setApplicationName(APP); app.setWindowIcon(qicon()); app.setQuitOnLastWindowClosed(False)
    apply_dark(app); app.setStyleSheet(settings_qss())
    log("Qt запущен, ищу уже работающую копию этой версии...")
    if send_to_running(args): log("Копия этой версии уже работает - команда передана ей, выхожу"); return
    if not args["data"]:
        kill_other_instances(); time.sleep(0.6); log("Старые копии закрыты (если были)")
    if IS_WIN and not args["data"]:
        for fn, key in ((set_ctx, "ctx"), (set_autostart, "autostart")):
            try: fn(S[key])
            except Exception: pass
    ctrl = Controller(app, args)
    log(f"трей: {ctrl.tray_ok}, горячие клавиши: {ctrl.hk_on} ({getattr(ctrl, 'hk_method', '')})")
    if args["files"] or args["convert"] or args["batch"] or args["single"]:
        act = "convert" if args["convert"] else "batch" if args["batch"] else "single" if args["single"] else "files"
        ctrl.collect(act, [str(Path(f).resolve()) for f in args["files"]], args["to"])
    elif not (args["tray"] and ctrl.tray_ok):
        ctrl.win.center_on_screen(); ctrl.win.popup(at_cursor=False)
        log("окно показано:", ctrl.win.isVisible(), ctrl.win.x(), ctrl.win.y(), ctrl.win.width(), ctrl.win.height())
    if ctrl.tray_ok and not args["data"]:
        QTimer.singleShot(1200, lambda: ctrl.tray.showMessage(APP, "Готово. " + (f"{current_combos()['single']} - один файл, {current_combos()['batch']} - несколько." if ctrl.hk_on else "Горячие клавиши не активны (см. журнал в меню трея)."), QSystemTrayIcon.MessageIcon.Information, 4000))
    code = app.exec()
    try: ctrl.win.hide()
    except Exception: pass
    os._exit(code)

if __name__ == "__main__":
    main()
