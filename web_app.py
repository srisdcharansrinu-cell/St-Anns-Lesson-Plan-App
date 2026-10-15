
from flask import (
    Flask, request, render_template_string,
    send_file, jsonify
)
from io import BytesIO
from pathlib import Path
import re
import os

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

APP_TITLE = "Lesson Plan Question Generator"

HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#075e54">
<title>Lesson Plan Question Generator</title>
<link rel="manifest" href="/manifest.json">
<style>
* { box-sizing: border-box; }
body {
    margin: 0;
    font-family: Arial, sans-serif;
    background: #f1f5f9;
    color: #172033;
}
header {
    background: linear-gradient(120deg,#075e54,#128c7e);
    color: white;
    padding: 28px 16px;
    text-align: center;
}
header h1 { margin: 0 0 10px; font-size: 27px; }
header p { margin: 0; }
.container {
    max-width: 1000px;
    margin: 24px auto;
    padding: 0 14px;
}
.card {
    background: white;
    padding: 22px;
    margin-bottom: 20px;
    border-radius: 14px;
    box-shadow: 0 4px 16px #0000000c;
}
h2 { color: #075e54; margin-top: 0; }
label {
    display: block;
    font-weight: bold;
    margin: 14px 0 7px;
}
input, textarea, select {
    width: 100%;
    padding: 12px;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    font-size: 15px;
    background: white;
}
textarea { min-height: 220px; resize: vertical; }
button, .download {
    border: 0;
    border-radius: 8px;
    padding: 12px 18px;
    cursor: pointer;
    font-size: 15px;
    font-weight: bold;
    text-decoration: none;
    display: inline-block;
    margin: 8px 5px 0 0;
}
.primary { background: #087f70; color: white; }
.secondary { background: #e2f3ef; color: #075e54; }
.question {
    padding: 12px;
    background: #f8fafc;
    border-left: 4px solid #128c7e;
    border-radius: 5px;
    margin: 10px 0;
    line-height: 1.5;
}
.notice { color: #475569; font-size: 14px; }
footer { text-align: center; padding: 22px; color: #64748b; }
@media(max-width:600px) {
    header h1 { font-size: 22px; }
    .card { padding: 16px; }
    button, .download { width: 100%; text-align: center; }
}
</style>
</head>
<body>
<header>
    <h1>Lesson Plan Question Generator</h1>
    <p>Generate study questions from your uploaded lesson content</p>
</header>

<main class="container">
<section class="card">
    <h2>Enter Lesson Plan</h2>
    <form method="post" action="/generate" enctype="multipart/form-data">
        <label for="subject">Subject Name</label>
        <input id="subject" name="subject"
               value="{{ subject }}" placeholder="Example: Enterprise Systems">

        <label for="unit">Unit / Topic</label>
        <input id="unit" name="unit"
               value="{{ unit }}" placeholder="Example: Unit 1 - Overview">

        <label for="lesson_file">Upload Lesson File (Optional)</label>
        <input id="lesson_file" type="file" name="lesson_file"
               accept=".txt,.md,.docx">
        <p class="notice">Supported files: TXT, MD and DOCX.</p>

        <label for="content">Paste Lesson Content</label>
        <textarea id="content" name="content"
          placeholder="Paste your lesson plan or question-and-answer material here...">{{ content }}</textarea>

        <label for="count">Number of Questions</label>
        <select id="count" name="count">
            {% for n in [5,10,15,20] %}
            <option value="{{ n }}" {% if count == n %}selected{% endif %}>
                {{ n }} Questions
            </option>
            {% endfor %}
        </select>

        <button class="primary" type="submit">Generate Questions</button>
        <button class="secondary" type="reset">Clear Form</button>
    </form>
</section>

{% if questions %}
<section class="card">
    <h2>Generated Questions</h2>
    <p><b>Subject:</b> {{ subject or "Not specified" }}</p>
    <p><b>Unit / Topic:</b> {{ unit or "Not specified" }}</p>
    <p class="notice">
      Questions are created from the text you provided.
      Please review them before using them in examinations.
    </p>

    {% for q in questions %}
      <div class="question">{{ loop.index }}. {{ q }}</div>
    {% endfor %}

    <form method="post" action="/download/txt">
      <input type="hidden" name="subject" value="{{ subject }}">
      <input type="hidden" name="unit" value="{{ unit }}">
      <input type="hidden" name="questions" value="{{ questions|join('\n') }}">
      <button class="secondary" type="submit">Download TXT</button>
    </form>

    <form method="post" action="/download/pdf">
      <input type="hidden" name="subject" value="{{ subject }}">
      <input type="hidden" name="unit" value="{{ unit }}">
      <input type="hidden" name="questions" value="{{ questions|join('\n') }}">
      <button class="primary" type="submit">Download PDF</button>
    </form>
</section>
{% endif %}
</main>
<footer>Lesson Plan Question Generator</footer>
<script>
if ("serviceWorker" in navigator) {
  window.addEventListener("load", function () {
    navigator.serviceWorker.register("/service-worker.js")
      .catch(function () {});
  });
}
</script>
</body>
</html>
"""


def read_uploaded_file(file):
    if not file or not file.filename:
        return ""

    name = file.filename.lower()

    if name.endswith(".txt") or name.endswith(".md"):
        return file.read().decode("utf-8", errors="replace")

    if name.endswith(".docx"):
        try:
            from docx import Document
            document = Document(file)
            return "\n".join(p.text for p in document.paragraphs)
        except ImportError:
            raise ValueError(
                "DOCX support needs python-docx. Run: "
                "python -m pip install python-docx"
            )
        except Exception as exc:
            raise ValueError("Could not read the DOCX file.") from exc

    raise ValueError("Please upload a TXT, MD, or DOCX file.")


def make_questions(text, limit=10):
    """Create simple questions using only the supplied lesson text."""
    text = re.sub(r"\r", "\n", text)
    lines = [
        re.sub(r"^\s*(?:Q\s*\d*[\.\):\-]?|[-*•])\s*", "", line).strip()
        for line in text.splitlines()
    ]
    lines = [
        line for line in lines
        if len(line) >= 12
        and not re.fullmatch(r"[\W_]+", line)
    ]

    results = []
    seen = set()

    def add(question):
        key = question.casefold()
        if key not in seen and len(results) < limit:
            seen.add(key)
            results.append(question)

    # Reuse existing questions in uploaded Q&A material.
    for line in lines:
        match = re.match(
            r"^(?:Q(?:uestion)?\s*\d*[\.\):\-]\s*)(.+\??)$",
            line,
            flags=re.IGNORECASE
        )
        if match:
            add(match.group(1).strip().rstrip("?") + "?")

    # Turn suitable lesson headings and statements into questions.
    for line in lines:
        cleaned = line.strip(" .:;")
        if not cleaned:
            continue

        if re.match(r"^(UNIT|CHAPTER|LESSON)\s*[-:]?\s*\d+",
                    cleaned, re.IGNORECASE):
            continue

        if len(cleaned) > 220:
            cleaned = cleaned[:220].rsplit(" ", 1)[0]

        if re.match(r"^(explain|describe|discuss|define|list|state)\b",
                    cleaned, re.IGNORECASE):
            add(cleaned.rstrip(".!?") + "?")
        elif len(cleaned.split()) <= 12:
            add("Explain " + cleaned.rstrip(".!?") + ".")
        else:
            add("What is meant by: " + cleaned.rstrip(".!?") + "?")

        if len(results) >= limit:
            break

    return results


@app.route("/", methods=["GET"])
def home():
    return render_template_string(
        HTML,
        subject="",
        unit="",
        content="",
        count=10,
        questions=[]
    )


@app.route("/generate", methods=["POST"])
def generate():
    subject = request.form.get("subject", "").strip()
    unit = request.form.get("unit", "").strip()
    content = request.form.get("content", "").strip()

    try:
        uploaded = read_uploaded_file(request.files.get("lesson_file"))
        if uploaded:
            content = (content + "\n" + uploaded).strip()
    except ValueError as exc:
        return str(exc), 400

    try:
        count = int(request.form.get("count", "10"))
    except ValueError:
        count = 10

    count = max(1, min(count, 20))

    if not content:
        return (
            "Please paste lesson content or upload a supported lesson file. "
            '<p><a href="/">Go back</a></p>',
            400
        )

    questions = make_questions(content, count)

    if not questions:
        questions = [
            "No suitable questions could be generated from this text. "
            "Please provide more lesson details."
        ]

    return render_template_string(
        HTML,
        subject=subject,
        unit=unit,
        content=content,
        count=count,
        questions=questions
    )


@app.route("/download/txt", methods=["POST"])
def download_txt():
    subject = request.form.get("subject", "")
    unit = request.form.get("unit", "")
    questions = request.form.get("questions", "")

    text = (
        "LESSON PLAN QUESTION GENERATOR\n\n"
        f"Subject: {subject}\n"
        f"Unit / Topic: {unit}\n\n"
        f"{questions}\n"
    )

    data = BytesIO(text.encode("utf-8"))
    data.seek(0)
    return send_file(
        data,
        mimetype="text/plain; charset=utf-8",
        as_attachment=True,
        download_name="lesson_questions.txt"
    )


@app.route("/download/pdf", methods=["POST"])
def download_pdf():
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer
        )
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import inch
    except ImportError:
        return (
            "PDF package missing. Run: "
            "python -m pip install reportlab",
            500
        )

    subject = request.form.get("subject", "")
    unit = request.form.get("unit", "")
    questions = request.form.get("questions", "")

    from xml.sax.saxutils import escape

    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=0.7 * inch,
        leftMargin=0.7 * inch,
        topMargin=0.7 * inch,
        bottomMargin=0.7 * inch
    )
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Lesson Plan Question Generator", styles["Title"]),
        Spacer(1, 14),
        Paragraph("<b>Subject:</b> " + escape(subject), styles["Normal"]),
        Paragraph("<b>Unit / Topic:</b> " + escape(unit), styles["Normal"]),
        Spacer(1, 16)
    ]

    for index, question in enumerate(questions.splitlines(), 1):
        if question.strip():
            story.append(
                Paragraph(
                    f"{index}. {escape(question)}",
                    styles["Normal"]
                )
            )
            story.append(Spacer(1, 9))

    doc.build(story)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="lesson_questions.pdf"
    )


@app.route("/manifest.json")
def manifest():
    return jsonify({
        "name": "Lesson Plan Question Generator",
        "short_name": "Question Generator",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#f1f5f9",
        "theme_color": "#075e54"
    })


@app.route("/service-worker.js")
def service_worker():
    code = """
self.addEventListener('install', event => self.skipWaiting());
self.addEventListener('activate', event => self.clients.claim());
"""
    from flask import Response
    return Response(code, mimetype="application/javascript")


if __name__ == "__main__":
    print("=" * 58)
    print("LESSON PLAN QUESTION GENERATOR")
    print("Open on this computer: http://127.0.0.1:5000")
    print("Keep this window open while using the application.")
    print("=" * 58)
    app.run(host="0.0.0.0", port=5000, debug=False)
