
import os
import io
import smtplib
import argparse
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
import matplotlib.pyplot as plt
from jinja2 import Environment, FileSystemLoader, select_autoescape
from dotenv import load_dotenv
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import schedule
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def load_env():
    # Load .env if present
    env_path = os.path.join(BASE_DIR, ".env")
    if os.path.exists(env_path):
        load_dotenv(env_path)

def read_data():
    data_path = os.path.join(BASE_DIR, "data", "sample_sales.csv")
    df = pd.read_csv(data_path, parse_dates=["date"])
    return df

def make_chart(df, out_path):
    # Simple revenue line
    plt.figure()
    plt.plot(df["date"], df["revenue"])
    plt.title("Revenue Over Time")
    plt.xlabel("Date")
    plt.ylabel("Revenue")
    plt.tight_layout()
    plt.savefig(out_path)
    plt.close()

def render_html(df, chart_path):
    env = Environment(
        loader=FileSystemLoader(os.path.join(BASE_DIR, "templates")),
        autoescape=select_autoescape()
    )
    template = env.get_template("report.html.j2")
    title = os.getenv("REPORT_TITLE", "Sales Report")
    total_orders = int(df["orders"].sum())
    total_revenue = float(df["revenue"].sum())
    best_row = df.sort_values("revenue", ascending=False).iloc[0]
    html = template.render(
        title=title,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        chart_path=os.path.basename(chart_path),
        total_orders=total_orders,
        total_revenue=f"${total_revenue:,.2f}",
        best_day={"date": best_row["date"].strftime("%Y-%m-%d"), "revenue": f"${best_row['revenue']:,.2f}"},
        rows=[{"date": d.strftime("%Y-%m-%d"), "orders": int(o), "revenue": f"${r:,.2f}"} for d, o, r in zip(df["date"], df["orders"], df["revenue"])]
    )
    return html

def export_pdf(chart_path, pdf_path, title):
    c = canvas.Canvas(pdf_path, pagesize=LETTER)
    width, height = LETTER
    y = height - 72
    c.setFont("Helvetica-Bold", 16)
    c.drawString(72, y, title)
    y -= 24
    c.setFont("Helvetica", 10)
    c.drawString(72, y, f"Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    y -= 24
    # Add chart
    with open(chart_path, "rb") as f:
        img = ImageReader(io.BytesIO(f.read()))
    c.drawImage(img, 72, y-300, width=width-144, height=300, preserveAspectRatio=True, anchor='n')
    c.showPage()
    c.save()

def send_email(subject, html_body, attachments):
    smtp_host = os.getenv("SMTP_HOST")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_pass = os.getenv("SMTP_PASS")
    mail_from = os.getenv("MAIL_FROM", smtp_user)
    mail_to = [addr.strip() for addr in os.getenv("MAIL_TO", smtp_user).split(",") if addr.strip()]

    if not all([smtp_host, smtp_user, smtp_pass]):
        raise RuntimeError("SMTP credentials missing. Please set SMTP_HOST/SMTP_USER/SMTP_PASS in .env")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = mail_from
    msg["To"] = ", ".join(mail_to)
    msg.set_content("Your email client does not support HTML.")
    msg.add_alternative(html_body, subtype="html")

    for fname, fbytes, mime in attachments:
        maintype, subtype = mime.split("/", 1)
        msg.add_attachment(fbytes, maintype=maintype, subtype=subtype, filename=fname)

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)

def build_and_send():
    df = read_data()
    out_dir = os.path.join(BASE_DIR, "out")
    os.makedirs(out_dir, exist_ok=True)

    chart_path = os.path.join(out_dir, "revenue_chart.png")
    make_chart(df, chart_path)

    html = render_html(df, chart_path)
    html_path = os.path.join(out_dir, "report.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    attachments = []
    with open(chart_path, "rb") as f:
        attachments.append(("revenue_chart.png", f.read(), "image/png"))

    enable_pdf = os.getenv("ENABLE_PDF", "true").lower() == "true"
    if enable_pdf:
        pdf_path = os.path.join(out_dir, "report.pdf")
        export_pdf(chart_path, pdf_path, os.getenv("REPORT_TITLE", "Sales Report"))
        with open(pdf_path, "rb") as f:
            attachments.append(("report.pdf", f.read(), "application/pdf"))

    subject = os.getenv("REPORT_TITLE", "Sales Report")
    send_email(subject, html, attachments)

def main():
    load_env()
    parser = argparse.ArgumentParser(description="Automated Email Report Sender")
    parser.add_argument("--send-now", action="store_true", help="Generate and send report immediately")
    parser.add_argument("--schedule", action="store_true", help="Run scheduler loop")
    args = parser.parse_args()

    if args.send_now:
        build_and_send()
        print("Report generated and email sent.")
        return

    if args.schedule:
        t = os.getenv("SCHEDULE_TIME", "09:00")
        schedule.every().day.at(t).do(build_and_send)
        print(f"Scheduler running. Will send every day at {t}. Press Ctrl+C to stop.")
        try:
            while True:
                schedule.run_pending()
                time.sleep(30)
        except KeyboardInterrupt:
            print("Stopped.")
        return

    parser.print_help()

if __name__ == "__main__":
    main()
