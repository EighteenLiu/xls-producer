# -*- coding: utf-8 -*-
"""Local web UI for generating the monthly street summary workbook.

Run:
    python app.py

Then open:
    http://127.0.0.1:8000
"""

from __future__ import annotations

import cgi
import json
import mimetypes
import shutil
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

from generate_monthly_summary import generate_summary, output_name_for, parse_ui_date
from report_generator import generate_reports


BASE_DIR = Path(__file__).resolve().parent
WORK_DIR = BASE_DIR / "web_work"
UPLOAD_DIR = WORK_DIR / "uploads"
DEFAULT_OUTPUT_DIR = WORK_DIR / "outputs"
HOST = "127.0.0.1"
PORT = 8000

DOWNLOAD_FILES: dict[str, Path] = {}


INDEX_HTML = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>街道问题汇总表生成</title>
  <style>
    :root {
      --bg: #f6f7f9;
      --panel: #ffffff;
      --text: #20242a;
      --muted: #667085;
      --line: #d9dee7;
      --accent: #1f7a6d;
      --accent-dark: #176258;
      --warn: #9a5b12;
      --error: #b42318;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Microsoft YaHei", "Segoe UI", Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
    }
    main {
      width: min(1040px, calc(100vw - 32px));
      margin: 28px auto 40px;
    }
    header {
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 20px;
      margin-bottom: 18px;
    }
    h1 {
      margin: 0;
      font-size: 26px;
      letter-spacing: 0;
      font-weight: 700;
    }
    .subtitle {
      margin: 8px 0 0;
      color: var(--muted);
      font-size: 14px;
    }
    .layout {
      display: grid;
      grid-template-columns: 1.05fr .95fr;
      gap: 18px;
    }
    section, aside {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 20px;
    }
    .grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
    }
    label {
      display: block;
      font-size: 13px;
      color: #344054;
      margin-bottom: 8px;
      font-weight: 600;
    }
    input[type="file"], input[type="date"], input[type="text"], select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      color: var(--text);
      padding: 11px 12px;
      font-size: 14px;
    }
    input[type="file"] { min-height: 44px; }
    .field { margin-bottom: 16px; }
    .path-row {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 10px;
      align-items: center;
    }
    .hint {
      margin-top: 6px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.5;
    }
    .actions {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-top: 4px;
    }
    button, .download {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      min-height: 42px;
      border: 0;
      border-radius: 6px;
      padding: 0 18px;
      background: var(--accent);
      color: #fff;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      white-space: nowrap;
    }
    button:hover, .download:hover { background: var(--accent-dark); }
    .secondary {
      min-height: 42px;
      background: #eef4f3;
      color: var(--accent-dark);
      border: 1px solid #b7d4ce;
    }
    .secondary:hover { background: #deebe8; }
    button:disabled {
      opacity: .6;
      cursor: not-allowed;
    }
    .output-name {
      padding: 12px;
      border-radius: 6px;
      border: 1px dashed #b8c4d6;
      background: #fbfcfe;
      font-family: Consolas, "Microsoft YaHei", monospace;
      font-size: 13px;
      overflow-wrap: anywhere;
    }
    .status {
      min-height: 42px;
      margin-top: 16px;
      padding: 12px;
      border-radius: 6px;
      background: #f2f4f7;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.6;
      white-space: pre-wrap;
    }
    .status.ok { color: #075e54; background: #eaf6f2; }
    .status.warn { color: var(--warn); background: #fff7e8; }
    .status.error { color: var(--error); background: #fff0ee; }
    .summary-title {
      margin: 0 0 12px;
      font-size: 16px;
    }
    .kv {
      display: grid;
      grid-template-columns: 112px 1fr;
      gap: 8px 12px;
      font-size: 14px;
      margin-bottom: 18px;
    }
    .kv span:nth-child(odd) { color: var(--muted); }
    .counts {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 8px 12px;
      font-size: 14px;
      padding-top: 12px;
      border-top: 1px solid var(--line);
    }
    .missing {
      margin: 12px 0 0;
      padding-left: 18px;
      color: var(--warn);
      font-size: 13px;
      line-height: 1.6;
    }
    .report-block {
      grid-column: 1;
      margin-top: 18px;
    }
    .layout > aside {
      grid-column: 2;
      grid-row: 1 / span 2;
    }
    .download-list {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 12px;
    }
    .notes {
      margin-top: 18px;
      padding-top: 14px;
      border-top: 1px solid var(--line);
      color: var(--muted);
      font-size: 13px;
      line-height: 1.7;
    }
    .notes p {
      margin: 0 0 8px;
    }
    @media (max-width: 760px) {
      header, .layout, .grid { display: block; }
      .layout > aside { display: block; }
      aside { margin-top: 18px; }
      .path-row { grid-template-columns: 1fr; }
      .actions { align-items: stretch; flex-direction: column; }
      button, .download { width: 100%; }
    }
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>街道问题汇总表生成</h1>
        <p class="subtitle">上传汇总 xlsx 模板和数据源台账，选择日期区间和保存地址后生成汇总表。</p>
      </div>
    </header>
    <div class="layout">
      <section>
        <form id="form">
          <div class="field">
            <label for="template">汇总 xlsx 模板</label>
            <input id="template" name="template" type="file" accept=".xlsx" required />
          </div>
          <div class="field">
            <label for="source">数据源台账</label>
            <input id="source" name="source" type="file" accept=".xlsx" required />
          </div>
          <div class="grid">
            <div class="field">
              <label for="start_date">开始日期</label>
              <input id="start_date" name="start_date" type="date" value="2026-04-20" required />
            </div>
            <div class="field">
              <label for="end_date">结束日期</label>
              <input id="end_date" name="end_date" type="date" value="2026-05-19" required />
            </div>
          </div>
          <div class="field">
            <label for="output_dir">文件保存地址</label>
            <div class="path-row">
              <input id="output_dir" name="output_dir" type="text" value="D:/pycharm/xls-producer/web_work/outputs" required />
              <button id="chooseOutputDir" class="secondary" type="button">选择文件夹</button>
            </div>
            <div class="hint">生成文件会直接保存到该文件夹，也可以手动粘贴路径。</div>
          </div>
          <div class="field">
            <label>生成文件名</label>
            <div class="output-name" id="outputName">（汇总）2026.5月各街道问题汇总.xlsx</div>
          </div>
          <div class="actions">
            <button id="submit" type="submit">生成汇总表</button>
            <a id="download" class="download" href="#" hidden>下载结果</a>
          </div>
          <div id="status" class="status">等待上传文件。</div>
        </form>
      </section>
      <section class="report-block">
        <form id="reportForm">
          <div class="field">
            <label for="summary_xlsx">各街道问题汇总表</label>
            <input id="summary_xlsx" name="summary_xlsx" type="file" accept=".xlsx" required />
          </div>
          <div class="field">
            <label for="source_ledger">源台账</label>
            <input id="source_ledger" name="source_ledger" type="file" accept=".xlsx" required />
          </div>
          <div class="grid">
            <div class="field">
              <label for="residential_template">垃圾分类检查自查报告模板</label>
              <input id="residential_template" name="residential_template" type="file" accept=".docx" required />
            </div>
            <div class="field">
              <label for="social_template">社会单位、餐饮单位检查报告模板</label>
              <input id="social_template" name="social_template" type="file" accept=".docx" required />
            </div>
          </div>
          <div class="grid">
            <div class="field">
              <label for="report_start_date">开始日期</label>
              <input id="report_start_date" name="start_date" type="date" value="2026-04-20" required />
            </div>
            <div class="field">
              <label for="report_end_date">结束日期</label>
              <input id="report_end_date" name="end_date" type="date" value="2026-05-19" required />
            </div>
          </div>
          <div class="field">
            <label for="report_output_dir">报告保存地址</label>
            <div class="path-row">
              <input id="report_output_dir" name="output_dir" type="text" value="D:/pycharm/xls-producer/web_work/outputs" required />
              <button id="chooseReportOutputDir" class="secondary" type="button">选择文件夹</button>
            </div>
          </div>
          <div class="actions">
            <button id="reportSubmit" type="submit">生成两个报告</button>
          </div>
          <div id="reportStatus" class="status">等待上传报告模板。</div>
          <div id="reportDownloads" class="download-list"></div>
        </form>
      </section>
      <aside>
        <h2 class="summary-title">生成结果</h2>
        <div class="kv">
          <span>时间跨度</span><strong id="rangeText">2026-04-20 至 2026-05-19</strong>
          <span>月份依据</span><strong>结束日期</strong>
          <span>输出月份</span><strong id="monthText">2026 年 5 月</strong>
          <span>保存位置</span><strong id="saveText">D:/pycharm/xls-producer/web_work/outputs</strong>
        </div>
        <div class="counts" id="counts">
          <span>暂无数据</span><strong>-</strong>
        </div>
        <ul class="missing" id="missing"></ul>
        <div class="notes">
          <p><strong>使用须知</strong></p>
          <p>先生成汇总表，再用汇总表和源台账生成两个 Word 报告。</p>
          <p>模板和台账请保持工作表名称、街道列和日期列结构稳定。</p>
          <p>如果目标文件已打开，生成时可能无法覆盖，请先关闭同名文件。</p>
        </div>
      </aside>
    </div>
  </main>
  <script>
    const form = document.getElementById("form");
    const startInput = document.getElementById("start_date");
    const endInput = document.getElementById("end_date");
    const outputDirInput = document.getElementById("output_dir");
    const chooseOutputDir = document.getElementById("chooseOutputDir");
    const outputName = document.getElementById("outputName");
    const rangeText = document.getElementById("rangeText");
    const monthText = document.getElementById("monthText");
    const saveText = document.getElementById("saveText");
    const statusBox = document.getElementById("status");
    const submit = document.getElementById("submit");
    const download = document.getElementById("download");
    const counts = document.getElementById("counts");
    const missing = document.getElementById("missing");
    const reportForm = document.getElementById("reportForm");
    const reportSubmit = document.getElementById("reportSubmit");
    const reportStatus = document.getElementById("reportStatus");
    const reportDownloads = document.getElementById("reportDownloads");
    const reportStartInput = document.getElementById("report_start_date");
    const reportEndInput = document.getElementById("report_end_date");
    const reportOutputDirInput = document.getElementById("report_output_dir");
    const chooseReportOutputDir = document.getElementById("chooseReportOutputDir");

    function updatePreview() {
      const start = startInput.value;
      const end = endInput.value;
      if (!end) return;
      const date = new Date(end + "T00:00:00");
      const year = date.getFullYear();
      const month = date.getMonth() + 1;
      outputName.textContent = `（汇总）${year}.${month}月各街道问题汇总.xlsx`;
      rangeText.textContent = `${start || "-"} 至 ${end}`;
      monthText.textContent = `${year} 年 ${month} 月`;
      saveText.textContent = outputDirInput.value || "-";
    }

    function setStatus(text, cls) {
      statusBox.className = "status" + (cls ? " " + cls : "");
      statusBox.textContent = text;
    }

    startInput.addEventListener("change", updatePreview);
    endInput.addEventListener("change", updatePreview);
    outputDirInput.addEventListener("input", updatePreview);
    updatePreview();

    async function chooseDirectory(input, afterChange) {
      const current = encodeURIComponent(input.value || "");
      const response = await fetch(`/choose-output-dir?current=${current}`);
      const result = await response.json();
      if (!response.ok || !result.ok) {
        throw new Error(result.error || "选择文件夹失败");
      }
      if (result.path) {
        input.value = result.path;
        if (afterChange) afterChange();
      }
    }

    chooseOutputDir.addEventListener("click", async () => {
      chooseOutputDir.disabled = true;
      try {
        await chooseDirectory(outputDirInput, updatePreview);
      } catch (error) {
        setStatus(error.message, "error");
      } finally {
        chooseOutputDir.disabled = false;
      }
    });

    chooseReportOutputDir.addEventListener("click", async () => {
      chooseReportOutputDir.disabled = true;
      try {
        await chooseDirectory(reportOutputDirInput);
      } catch (error) {
        reportStatus.className = "status error";
        reportStatus.textContent = error.message;
      } finally {
        chooseReportOutputDir.disabled = false;
      }
    });

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      download.hidden = true;
      missing.innerHTML = "";

      if (startInput.value > endInput.value) {
        setStatus("开始日期不能晚于结束日期。", "error");
        return;
      }

      submit.disabled = true;
      setStatus("正在上传并生成，请稍候...", "");
      try {
        const body = new FormData(form);
        const response = await fetch("/generate", { method: "POST", body });
        const result = await response.json();
        if (!response.ok || !result.ok) {
          throw new Error(result.error || "生成失败");
        }

        const message = `生成完成：${result.filename}\n保存位置：${result.output_path}`;
        setStatus(message, result.missing.length ? "warn" : "ok");
        download.href = result.download_url;
        download.download = result.filename;
        download.hidden = false;

        counts.innerHTML = "";
        for (const [name, count] of Object.entries(result.row_counts)) {
          counts.insertAdjacentHTML("beforeend", `<span>${name}</span><strong>${count}</strong>`);
        }
        missing.innerHTML = result.missing.map(item => `<li>${item}</li>`).join("");
      } catch (error) {
        setStatus(error.message, "error");
      } finally {
        submit.disabled = false;
      }
    });

    reportForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      reportDownloads.innerHTML = "";

      if (reportStartInput.value > reportEndInput.value) {
        reportStatus.className = "status error";
        reportStatus.textContent = "开始日期不能晚于结束日期。";
        return;
      }

      reportSubmit.disabled = true;
      reportStatus.className = "status";
      reportStatus.textContent = "正在上传并生成两个报告，请稍候...";
      try {
        const body = new FormData(reportForm);
        const response = await fetch("/generate-reports", { method: "POST", body });
        const result = await response.json();
        if (!response.ok || !result.ok) {
          throw new Error(result.error || "报告生成失败");
        }

        reportStatus.className = "status" + (result.missing.length ? " warn" : " ok");
        reportStatus.textContent = `生成完成：${result.files.length} 个报告\n保存位置：${result.output_dir}`;
        reportDownloads.innerHTML = result.files.map(file =>
          `<a class="download" href="${file.download_url}" download="${file.filename}">${file.filename}</a>`
        ).join("");
      } catch (error) {
        reportStatus.className = "status error";
        reportStatus.textContent = error.message;
      } finally {
        reportSubmit.disabled = false;
      }
    });
  </script>
</body>
</html>
"""


def json_response(handler: BaseHTTPRequestHandler, status: int, payload: dict) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def save_upload(field, directory: Path, fallback_name: str, allowed_extensions: tuple[str, ...] = (".xlsx",)) -> Path:
    filename = Path(field.filename or fallback_name).name
    if not filename.lower().endswith(allowed_extensions):
        allowed = "、".join(allowed_extensions)
        raise ValueError(f"{filename} 不是 {allowed} 文件")
    stamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
    target = directory / f"{stamp}_{filename}"
    with target.open("wb") as output:
        shutil.copyfileobj(field.file, output)
    return target


def resolve_output_dir(value: str | None) -> Path:
    text = (value or "").strip().strip('"')
    path = Path(text) if text else DEFAULT_OUTPUT_DIR
    if not path.is_absolute():
        path = (BASE_DIR / path).resolve()
    path.mkdir(parents=True, exist_ok=True)
    if not path.is_dir():
        raise ValueError(f"{path} 不是有效文件夹")
    return path


def register_download(path: Path) -> str:
    key = f"{datetime.now():%Y%m%d%H%M%S%f}_{path.name}"
    DOWNLOAD_FILES[key] = path
    return f"/download/{quote(key)}"


def choose_output_directory(current: str | None = None) -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:
        raise RuntimeError("当前环境无法打开文件夹选择窗口，请手动填写保存路径。") from exc

    initial = (current or "").strip().strip('"')
    if not initial or not Path(initial).exists():
        initial = str(DEFAULT_OUTPUT_DIR)

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askdirectory(
            title="选择保存文件夹",
            initialdir=initial,
            mustexist=False,
            parent=root,
        )
    finally:
        root.destroy()
    return selected


class AppHandler(BaseHTTPRequestHandler):
    server_version = "XlsProducerUI/2.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = INDEX_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path == "/choose-output-dir":
            self.choose_output_dir(parsed)
            return
        if parsed.path.startswith("/download/"):
            self.serve_download(parsed.path)
            return
        self.send_error(404)

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path == "/generate-reports":
            self.generate_report_docs()
            return
        if path != "/generate":
            self.send_error(404)
            return
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                    "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
                },
            )
            start_date = parse_ui_date(form.getfirst("start_date", ""))
            end_date = parse_ui_date(form.getfirst("end_date", ""), end_of_day=True)
            if start_date > end_date:
                raise ValueError("开始日期不能晚于结束日期")

            UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_dir = resolve_output_dir(form.getfirst("output_dir", ""))
            template_file = save_upload(form["template"], UPLOAD_DIR, "template.xlsx")
            source_file = save_upload(form["source"], UPLOAD_DIR, "source.xlsx")
            filename = output_name_for(end_date)
            output_file = output_dir / filename

            result = generate_summary(source_file, template_file, start_date, end_date, output_file)
            json_response(
                self,
                200,
                {
                    "ok": True,
                    "filename": filename,
                    "download_url": register_download(output_file),
                    "output_path": str(output_file),
                    "row_counts": result["row_counts"],
                    "missing": result["missing"],
                },
            )
        except Exception as exc:
            json_response(self, 400, {"ok": False, "error": str(exc)})

    def generate_report_docs(self) -> None:
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                    "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
                },
            )
            start_date = parse_ui_date(form.getfirst("start_date", ""))
            end_date = parse_ui_date(form.getfirst("end_date", ""), end_of_day=True)
            if start_date > end_date:
                raise ValueError("开始日期不能晚于结束日期")

            UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_dir = resolve_output_dir(form.getfirst("output_dir", ""))
            summary_file = save_upload(form["summary_xlsx"], UPLOAD_DIR, "summary.xlsx")
            source_file = save_upload(form["source_ledger"], UPLOAD_DIR, "source.xlsx")
            residential_template = save_upload(
                form["residential_template"],
                UPLOAD_DIR,
                "residential_template.docx",
                (".docx",),
            )
            social_template = save_upload(
                form["social_template"],
                UPLOAD_DIR,
                "social_template.docx",
                (".docx",),
            )

            result = generate_reports(
                summary_file,
                source_file,
                residential_template,
                social_template,
                start_date,
                end_date,
                output_dir,
            )
            files = [
                {
                    "filename": path.name,
                    "download_url": register_download(path),
                    "output_path": str(path),
                }
                for path in result["output_files"]
            ]
            json_response(
                self,
                200,
                {
                    "ok": True,
                    "files": files,
                    "output_dir": str(output_dir),
                    "missing": result["missing"],
                },
            )
        except Exception as exc:
            json_response(self, 400, {"ok": False, "error": str(exc)})

    def serve_download(self, path_text: str) -> None:
        key = Path(unquote(path_text.removeprefix("/download/"))).name
        path = DOWNLOAD_FILES.get(key)
        if path is None or not path.exists():
            self.send_error(404, "File not found")
            return
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(path.name)}")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def choose_output_dir(self, parsed) -> None:
        try:
            current = parse_qs(parsed.query).get("current", [""])[0]
            path = choose_output_directory(current)
            json_response(self, 200, {"ok": True, "path": path})
        except Exception as exc:
            json_response(self, 400, {"ok": False, "error": str(exc)})

    def log_message(self, format: str, *args) -> None:
        print(f"{self.address_string()} - {format % args}")


class ReusableThreadingHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True


def main() -> None:
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    server = ReusableThreadingHTTPServer((HOST, PORT), AppHandler)
    print(f"街道问题汇总表生成 UI: http://{HOST}:{PORT}")
    print("按 Ctrl+C 停止服务。")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        threading.Thread(target=server.server_close).start()


if __name__ == "__main__":
    main()
