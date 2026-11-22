"""生成用于验证解析与定位的多种格式样本。"""
from __future__ import annotations

import os
import sqlite3
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "samples")


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    secrets = ["张三", "410105199001011234", "secret_token_9f3a", "192.168.1.88", "admin@example.com"]

    # 1. 纯文本
    with open(os.path.join(OUT, "notes.txt"), "w", encoding="utf-8") as f:
        f.write("会议记录\n")
        f.write("参与人：张三、李四\n")
        f.write("身份证号 410105199001011234 已登记。\n")
        f.write("服务器地址 192.168.1.88，密码 secret_token_9f3a\n")

    # 2. GBK 文本（编码探测）
    with open(os.path.join(OUT, "legacy_gbk.txt"), "wb") as f:
        f.write("旧系统导出：张三的联系方式 admin@example.com\n".encode("gbk"))

    # 3. CSV
    with open(os.path.join(OUT, "data.csv"), "w", encoding="utf-8") as f:
        f.write("姓名,身份证,邮箱\n张三,410105199001011234,admin@example.com\n李四,410105199202022345,li@example.com\n")

    # 4. JSON
    with open(os.path.join(OUT, "config.json"), "w", encoding="utf-8") as f:
        f.write('{"host": "192.168.1.88", "token": "secret_token_9f3a", "owner": "张三"}\n')

    # 5. DOCX
    try:
        import docx
        d = docx.Document()
        d.add_heading("项目报告", 0)
        d.add_paragraph("负责人：张三，联系邮箱 admin@example.com")
        d.add_paragraph("服务器 192.168.1.88 需要迁移。")
        t = d.add_table(rows=2, cols=2)
        t.cell(0, 0).text = "编号"
        t.cell(0, 1).text = "身份证"
        t.cell(1, 0).text = "001"
        t.cell(1, 1).text = "410105199001011234"
        d.save(os.path.join(OUT, "报告.docx"))
    except Exception as e:
        print("docx skip:", e)

    # 6. XLSX
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "人员"
        ws.append(["姓名", "身份证", "备注"])
        ws.append(["张三", "410105199001011234", "secret_token_9f3a"])
        ws2 = wb.create_sheet("服务器")
        ws2.append(["主机", "IP"])
        ws2.append(["web01", "192.168.1.88"])
        wb.save(os.path.join(OUT, "账本.xlsx"))
    except Exception as e:
        print("xlsx skip:", e)

    # 7. PPTX
    try:
        from pptx import Presentation
        from pptx.util import Inches
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "安全审计"
        slide.placeholders[1].text = "发现 admin@example.com 异常登录\nIP: 192.168.1.88"
        prs.save(os.path.join(OUT, "演示.pptx"))
    except Exception as e:
        print("pptx skip:", e)

    # 8. PDF
    try:
        import pymupdf
        doc = pymupdf.open()
        page = doc.new_page()
        page.insert_text((72, 100), "Internal Audit Report", fontsize=18)
        page.insert_text((72, 130), "Contact: admin@example.com", fontsize=12)
        page.insert_text((72, 150), "IP address: 192.168.1.88", fontsize=12)
        page2 = doc.new_page()
        page2.insert_text((72, 100), "Owner: 张三 (Zhang San)", fontsize=12)
        page2.insert_text((72, 120), "Token: secret_token_9f3a", fontsize=12)
        doc.save(os.path.join(OUT, "manual.pdf"))
        doc.close()
    except Exception as e:
        print("pdf skip:", e)

    # 9. SQLite
    db = os.path.join(OUT, "users.db")
    if os.path.exists(db):
        os.unlink(db)
    con = sqlite3.connect(db)
    con.execute("CREATE TABLE users(id INTEGER PRIMARY KEY, name TEXT, idcard TEXT, email TEXT)")
    con.execute("CREATE TABLE servers(host TEXT, ip TEXT, note TEXT)")
    con.executemany("INSERT INTO users(name,idcard,email) VALUES(?,?,?)", [
        ("张三", "410105199001011234", "admin@example.com"),
        ("李四", "410105199202022345", "li@example.com"),
    ])
    con.execute("INSERT INTO servers VALUES('web01','192.168.1.88','secret_token_9f3a')")
    con.commit()
    con.close()

    # 10. EML
    with open(os.path.join(OUT, "mail.eml"), "w", encoding="utf-8") as f:
        f.write("From: a@example.com\nTo: b@example.com\nSubject: 账号密码\n"
                "Date: Mon, 1 Jan 2024 10:00:00 +0800\n"
                "Content-Type: text/plain; charset=utf-8\n\n"
                "张三的账号密码是 secret_token_9f3a，服务器 192.168.1.88。\n")

    # 11. ZIP（内含 txt + docx）
    with zipfile.ZipFile(os.path.join(OUT, "bundle.zip"), "w") as z:
        z.writestr("inner/readme.txt", "压缩包内部：张三 admin@example.com\n")
        if os.path.exists(os.path.join(OUT, "报告.docx")):
            z.write(os.path.join(OUT, "报告.docx"), "inner/报告.docx")

    # 12. RTF
    with open(os.path.join(OUT, "legacy.rtf"), "w", encoding="utf-8") as f:
        f.write(r"{\rtf1\ansi\deff0 {\fonttbl{\f0 Helvetica;}}\f0\fs24 Legacy RTF: secret_token_9f3a\par}")

    # 13. 大日志（性能）
    log_path = os.path.join(OUT, "big.log")
    if not os.path.exists(log_path):
        with open(log_path, "w", encoding="utf-8") as f:
            for i in range(20000):
                f.write(f"2024-01-01 00:00:{i%60:02d} INFO request id={i} user=u{i%50} ip=10.0.0.{i%255}\n")
            f.write("2024-01-01 23:59:59 ERROR login failed for 张三 from 192.168.1.88\n")

    # 14. 二进制（模拟镜像，含 ASCII 串）
    with open(os.path.join(OUT, "raw.bin"), "wb") as f:
        f.write(b"\x00\x01\x02" * 1000 + b"\x00" + b"password=secret_token_9f3a\x00" + b"\x7f" * 500)

    total = sum(os.path.getsize(os.path.join(OUT, n)) for n in os.listdir(OUT))
    print(f"生成 {len(os.listdir(OUT))} 个样本 -> {OUT} ({total/1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
