import os, sys, subprocess, ssl
from email.message import EmailMessage
import smtplib
from docxtpl import DocxTemplate
import pandas as pd
from pathlib import Path
import configparser
import time

ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))  # PyInstaller 대응
APPDIR = Path(sys.argv[0]).resolve().parent

DATA_DIR = APPDIR / "data"
OUT_DIR = APPDIR / "output"
CONF_PATH = APPDIR / "config.ini"

def load_config():
    cfg = configparser.ConfigParser()
    if CONF_PATH.exists():
        cfg.read(CONF_PATH, encoding="utf-8")
    else:
        cfg["SMTP"] = {
            "HOST": os.getenv("SMTP_HOST", "smtp.daum.net"),
            "PORT": os.getenv("SMTP_PORT", "465"),
            "USER": os.getenv("SMTP_USER", ""),
            "PASS": os.getenv("SMTP_PASS", ""),
            "USE_SSL": os.getenv("SMTP_USE_SSL", "true"),
            "DRY_RUN": os.getenv("SMTP_DRY_RUN", "false"),
            "SKIP_EMAIL": os.getenv("SMTP_SKIP_EMAIL", "false")
        }
        cfg["CONVERT"] = {
            "LIBREOFFICE_EXE": os.getenv("LIBREOFFICE_EXE", str(APPDIR / "LibreOfficePortable" / "App" / "libreoffice" / "program" / "soffice.exe")),
            "SKIP_CONVERT": os.getenv("CONVERT_SKIP_CONVERT", "false")
        }
        cfg["APP"] = {
            "SUBJECT_PREFIX": os.getenv("SUBJECT_PREFIX", "[안내] 문서 발송 - "),
            "SLEEP_BETWEEN_MS": os.getenv("SLEEP_BETWEEN_MS", "0")
        }
    return cfg

def docx_to_pdf(libreo_exe: Path, docx_path: Path, pdf_path: Path):
    outdir = pdf_path.parent
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(libreo_exe),
        "--headless", "--nologo", "--nolockcheck", "--nodefault",
        "--convert-to", "pdf",
        "--outdir", str(outdir),
        str(docx_path)
    ]
    subprocess.run(cmd, check=True)

def send_email(smtp_host, smtp_port, user, password, use_ssl, to_addr, subject, body, attachments):
    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)

    for path in attachments:
        with open(path, "rb") as f:
            data = f.read()
        fname = Path(path).name
        msg.add_attachment(data, maintype="application", subtype="pdf", filename=fname)

    if use_ssl:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_host, int(smtp_port), context=context) as s:
            s.login(user, password)
            s.send_message(msg)
    else:
        with smtplib.SMTP(smtp_host, int(smtp_port)) as s:
            s.starttls(context=ssl.create_default_context())
            s.login(user, password)
            s.send_message(msg)

def find_single_template():
    docs = sorted([p for p in DATA_DIR.glob("*.docx") if p.is_file()])
    if len(docs) == 0:
        print("ERROR: data/ 폴더에 .docx 템플릿이 없습니다.")
        sys.exit(1)
    if len(docs) > 1:
        print("ERROR: data/ 폴더에 .docx 파일이 2개 이상 있습니다. 하나만 남겨주세요:")
        for p in docs:
            print(" -", p.name)
        sys.exit(1)
    return docs[0]

def main():
    cfg = load_config()
    smtp = cfg["SMTP"]
    conv = cfg["CONVERT"]
    app = cfg["APP"] if "APP" in cfg else {"SUBJECT_PREFIX":"[안내] 문서 발송 - ", "SLEEP_BETWEEN_MS":"0"}

    # 1) 템플릿 자동 탐색
    template_path = find_single_template()

    csv_path = DATA_DIR / "data.csv"
    OUT_DIR.mkdir(exist_ok=True, parents=True)

    if not csv_path.exists():
        print("ERROR: data/data.csv 가 없습니다.")
        sys.exit(1)

    dry_run = str(smtp.get("DRY_RUN", "false")).lower() == "true"
    skip_email = str(smtp.get("SKIP_EMAIL", "false")).lower() == "true"
    skip_convert = str(conv.get("SKIP_CONVERT", "false")).lower() == "true"
    sleep_ms = int(app.get("SLEEP_BETWEEN_MS", "0"))

    df = pd.read_csv(csv_path)

    for idx, row in df.iterrows():
        ctx = row.to_dict()
        # DOCX 출력 이름 = 템플릿이름 + 순번 + 이름
        base_stem = template_path.stem
        suffix = f"_{idx+1}_{ctx.get('name','')}".strip('_')
        docx_out = OUT_DIR / f"{base_stem}{suffix}.docx"

        doc = DocxTemplate(str(template_path))
        doc.render(ctx)
        doc.save(str(docx_out))

        # PDF 출력 이름 = DOCX와 동일
        pdf_out = docx_out.with_suffix('.pdf')
        if dry_run or skip_convert:
            with open(pdf_out.with_suffix('.pdf.skip'), 'w', encoding='utf-8') as f:
                f.write("SKIP_CONVERT or DRY_RUN: PDF 변환 스킵")
        else:
            docx_to_pdf(Path(conv["LIBREOFFICE_EXE"]), docx_out, pdf_out)

        # 메일 발송
        to_addr = ctx["email"]
        subject = f"{app.get('SUBJECT_PREFIX','[안내] 문서 발송 - ')}{ctx.get('name','')}"
        body = f"""{ctx.get('name','')}님,

첨부된 PDF를 확인해 주세요.

감사합니다.
"""
        if dry_run or skip_email:
            print(f"[SKIP_EMAIL/DRY_RUN] Would send to {to_addr}: {pdf_out.name}")
        else:
            send_email(
                smtp.get("HOST","smtp.daum.net"), smtp.get("PORT","465"),
                smtp.get("USER",""), smtp.get("PASS",""),
                str(smtp.get("USE_SSL","true")).lower()=="true",
                to_addr, subject, body, [str(pdf_out)]
            )
            print(f"Sent to {to_addr}: {pdf_out.name}")

        if sleep_ms > 0:
            time.sleep(sleep_ms/1000.0)

    print("ALL DONE")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR:", e)
        sys.exit(1)
