"""
Agentic Test Case Synthesizer — core pipeline.

Orchestrates: ADO story fetch → LLM generation (+ optional review) → ADO upload.
All user-visible logs follow a consistent emoji/style contract that the UI
log-box uses for color coding (info=blue, ok=green, warn=yellow, err=red).
"""
from __future__ import annotations

import ast
import base64
import os
import queue
import re
from datetime import datetime
from html import unescape
from pathlib import Path
from typing import Any, Callable

import openai
import pandas as pd
import requests
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth

# ── Config ────────────────────────────────────────────────────────────────────
load_dotenv()
openai.api_type    = "azure"
openai.api_base    = os.getenv("AZURE_OPENAI_ENDPOINT")
openai.api_version = "2024-02-15-preview"
openai.api_key     = os.getenv("AZURE_OPENAI_KEY")
_DEPLOYMENT        = os.getenv("AZURE_OPENAI_DEPLOYMENT")

_missing = [k for k, v in {
    "AZURE_OPENAI_ENDPOINT":   openai.api_base,
    "AZURE_OPENAI_KEY":        openai.api_key,
    "AZURE_OPENAI_DEPLOYMENT": _DEPLOYMENT,
}.items() if not v]
if _missing:
    raise EnvironmentError(f"Missing env var(s): {', '.join(_missing)}")

_PROMPTS_DIR = Path(__file__).parent / "prompts"
_OUTPUT_DIR  = Path("output"); _OUTPUT_DIR.mkdir(exist_ok=True)

_TEXT_EXTS  = {".txt", ".md", ".markdown", ".json", ".xml", ".yaml", ".yml",
               ".csv", ".log", ".html", ".htm", ".rst", ".ini", ".toml"}
_IMAGE_MIME = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
               ".gif": "image/gif", ".webp": "image/webp"}
_IMAGE_EXTS = set(_IMAGE_MIME)


# ── Live log streaming ────────────────────────────────────────────────────────
# Log style guide (kept in sync with UI regex in test_case_synthesizer.html):
#   ⚡ 🔗 🧠 🔍 ☁️ 📋 📁 📂 📎 🖼   → info  (cyan)
#   ✅ ✔ "complete"                  → ok    (green)
#   ❌ "error" / "failed"             → err   (red)
#   ⚠️ "warn"                          → warn  (yellow)
# Indent sub-steps with two spaces. Keep lines short and scannable.

_log_queue: "queue.Queue | None" = None

def set_log_queue(q: "queue.Queue | None") -> None:
    global _log_queue
    _log_queue = q

def _log(msg: str = "") -> None:
    print(msg)
    if _log_queue is not None:
        _log_queue.put(str(msg))

def _hr() -> None:
    _log("─" * 50)


# ── Generic helpers ───────────────────────────────────────────────────────────
def _load_prompt(name: str) -> str:
    return (_PROMPTS_DIR / name).read_text(encoding="utf-8")

def _strip_html(text: Any) -> str:
    return re.sub(r"<[^>]+>", "", unescape(str(text or ""))).strip()

def _safe_filename(text: str, n: int = 50) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^\w\s\-]", "", text).strip())[:n].strip()

def _build_plan_suite(story_id: str, title: str,
                      plan_override: str | None, suite_override: str | None) -> tuple[str, str]:
    plan  = (plan_override  or "").strip() or f"TCS - #{story_id} - {_safe_filename(title)}"
    suite = (suite_override or "").strip() or "TCS"
    return plan, suite


# ── ADO HTTP helpers ──────────────────────────────────────────────────────────
def _ado_get(url: str, pat: str, *, timeout: int = 30, raw: bool = False):
    """GET against Azure DevOps. Returns parsed JSON, or raw response if raw=True."""
    r = requests.get(url, headers={"Content-Type": "application/json"},
                     auth=HTTPBasicAuth("", pat), timeout=timeout)
    return r if raw else (r.json() if r.status_code == 200 else None, r.status_code)

def _ado_post(url: str, pat: str, payload: dict, *, content_type: str = "application/json"):
    return requests.post(url, headers={"Content-Type": content_type},
                         auth=HTTPBasicAuth("", pat), json=payload)


# ── LLM ───────────────────────────────────────────────────────────────────────
def _llm(messages: list, *, temperature: float = 0.3) -> str:
    r = openai.ChatCompletion.create(engine=_DEPLOYMENT, messages=messages,
                                     temperature=temperature)
    return r.choices[0].message.content.strip()

def _build_user_content(prompt: str, images: list | None) -> Any:
    if not images:
        return prompt
    return [{"type": "text", "text": prompt}] + [
        {"type": "image_url",
         "image_url": {"url": f"data:{img['media_type']};base64,{img['data']}"}}
        for img in images
    ]


# ── ADO: story / comments / attachments ───────────────────────────────────────
def _fetch_comments(story_id: str, org: str, project: str, pat: str) -> str:
    url = (f"https://dev.azure.com/{org}/{project}/_apis/wit/workitems/"
           f"{story_id}/comments?api-version=7.1-preview.3")
    try:
        body, status = _ado_get(url, pat, timeout=15)
        if not body:
            _log(f"  ⚠️ Comments unavailable (HTTP {status}).")
            return ""
        parts = []
        for c in body.get("comments", []):
            text = _strip_html(c.get("text"))
            if text:
                author = c.get("createdBy", {}).get("displayName", "Unknown")
                date   = (c.get("createdDate") or "")[:10]
                parts.append(f"[{date}] {author}: {text}")
        if parts:
            _log(f"  💬 {len(parts)} comment(s) fetched.")
        return "\n".join(parts)
    except Exception as e:
        _log(f"  ⚠️ Comments error: {e}")
        return ""


def _fetch_attachments(relations: list, pat: str) -> tuple[str, list]:
    rels = [r for r in (relations or []) if r.get("rel") == "AttachedFile"]
    if not rels:
        return "", []
    _log(f"  📎 {len(rels)} attachment(s) found — downloading…")
    text_parts, images = [], []
    for rel in rels:
        url      = rel.get("url", "")
        filename = rel.get("attributes", {}).get("name", url.rsplit("/", 1)[-1])
        ext      = Path(filename).suffix.lower()
        try:
            resp = _ado_get(url, pat, timeout=60, raw=True)
            if resp.status_code != 200:
                _log(f"    ⚠️ {filename} skipped (HTTP {resp.status_code}).")
                text_parts.append(f"[Attachment: {filename} — could not download]")
                continue
            if ext in _TEXT_EXTS:
                text = resp.content.decode("utf-8", errors="replace").strip()
                text_parts.append(f"[Attachment: {filename}]\n{text}")
                _log(f"    ✔ Text · {filename} ({len(text)} chars)")
            elif ext in _IMAGE_EXTS:
                images.append({"media_type": _IMAGE_MIME[ext],
                               "data":       base64.b64encode(resp.content).decode(),
                               "name":       filename})
                _log(f"    🖼 Image · {filename}")
            else:
                text_parts.append(f"[Attachment: {filename} — binary, skipped]")
                _log(f"    · Binary · {filename}")
        except Exception as e:
            _log(f"    ⚠️ {filename} error: {e}")
            text_parts.append(f"[Attachment: {filename} — fetch error]")
    return "\n\n".join(text_parts), images


def _extract_inline_images(html: str, pat: str) -> list:
    images = []
    for url in re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', html or ""):
        if "dev.azure.com" not in url:
            continue
        try:
            resp = _ado_get(url, pat, timeout=30, raw=True)
            if resp.status_code != 200:
                continue
            ct  = resp.headers.get("content-type", "image/png").split(";")[0].strip()
            ext = "." + ct.split("/")[-1].replace("jpeg", "jpg")
            if ext not in _IMAGE_EXTS:
                continue
            images.append({"media_type": ct,
                           "data":       base64.b64encode(resp.content).decode(),
                           "name":       url.rsplit("/", 1)[-1]})
            _log(f"    🖼 Inline image fetched.")
        except Exception:
            pass
    return images


def _fetch_story(story_id: str, org: str, project: str, pat: str) -> dict:
    url = (f"https://dev.azure.com/{org}/{project}/_apis/wit/workitems/"
           f"{story_id}?$expand=all&api-version=7.0")
    body, status = _ado_get(url, pat)
    if not body:
        raise Exception(f"ADO story fetch failed (HTTP {status})")
    f        = body.get("fields", {})
    raw_desc = f.get("System.Description") or ""
    raw_ac   = f.get("Microsoft.VSTS.Common.AcceptanceCriteria") or ""

    inline = _extract_inline_images(raw_desc, pat) + _extract_inline_images(raw_ac, pat)
    attach_text, attach_imgs = _fetch_attachments(body.get("relations") or [], pat)
    images = inline + attach_imgs
    if images:
        _log(f"  🖼 {len(images)} image(s) collected total.")

    return {
        "id":                  story_id,
        "title":               (f.get("System.Title") or "").strip(),
        "description":         _strip_html(raw_desc),
        "acceptance_criteria": _strip_html(raw_ac),
        "comments":            _fetch_comments(story_id, org, project, pat),
        "attachments_text":    attach_text,
        "images":              images,
    }


def _fetch_suite_tcs(org: str, project: str, pat: str,
                     plan_id: str, suite_id: str) -> str:
    """Fetch all test cases from an ADO suite, formatted as plain text."""
    _log(f"  📂 Fetching reference suite — plan #{plan_id}, suite #{suite_id}…")
    base = f"https://dev.azure.com/{org}/{project}/_apis"

    body, status = _ado_get(
        f"{base}/testplan/Plans/{plan_id}/Suites/{suite_id}/TestCase?api-version=7.0", pat)
    if not body:
        raise Exception(f"Suite fetch failed (HTTP {status})")

    refs = body.get("value", [])
    if not refs:
        _log("  ⚠️ Reference suite is empty.")
        return ""

    ids = [str(r["workItem"]["id"]) for r in refs]
    _log(f"  📋 {len(ids)} reference test case(s) found.")

    detail, status = _ado_get(
        f"{base}/wit/workitems?ids={','.join(ids)}"
        f"&fields=System.Title,Microsoft.VSTS.TCM.Steps&api-version=7.0", pat)
    if not detail:
        raise Exception(f"Reference TC details fetch failed (HTTP {status})")

    parts = []
    for wi in detail.get("value", []):
        title = wi["fields"].get("System.Title", "Untitled")
        xml   = wi["fields"].get("Microsoft.VSTS.TCM.Steps") or ""
        raw   = re.findall(r"<parameterizedString[^>]*>(.*?)</parameterizedString>",
                           xml, re.DOTALL)
        lines = [
            f"  Action: {_strip_html(raw[i])}\n  Expected: {_strip_html(raw[i+1])}"
            for i in range(0, len(raw) - 1, 2) if _strip_html(raw[i])
        ]
        parts.append(f"Title: {title}\n" +
                     ("Steps:\n" + "\n".join(lines) if lines else "Steps: (none)"))

    _log("  ✔ Reference suite parsed.")
    return "\n\n---\n\n".join(parts)


# ── Agents ────────────────────────────────────────────────────────────────────
def _agent_generate(prompt_file: str, story: dict, suite_tcs: str = "") -> str:
    prompt = _load_prompt(prompt_file).format(
        user_story          = story["title"] or f"Story {story['id']}",
        acceptance_criteria = story["acceptance_criteria"] or "No acceptance criteria provided.",
        extra_context       = f"Description: {story['description']}" if story["description"] else "",
        comments            = story["comments"],
        attachments_text    = story["attachments_text"],
        reference_test_cases= suite_tcs,
        clone_test_cases    = suite_tcs,
        figma_instruction   = "",
    )
    images = story["images"]
    if images:
        _log(f"  🖼 {len(images)} image(s) → generation agent.")
    return _llm([
        {"role": "system", "content":
         "You are a QA engineer. Generate test cases in the EXACT format requested."},
        {"role": "user", "content": _build_user_content(prompt, images)},
    ])


def _agent_review(tcs: list, story: dict) -> list:
    if not tcs:
        return tcs
    numbered = "\n".join(
        f"{i}. Title: {tc['Title']}\n   Steps:\n{tc['Steps']}" for i, tc in enumerate(tcs, 1))
    prompt = _load_prompt("agent2_review.txt").format(
        test_cases          = numbered,
        user_story          = story["title"],
        acceptance_criteria = story["acceptance_criteria"],
    )
    images = story["images"]
    if images:
        _log(f"  🖼 {len(images)} image(s) → review agent.")
    response = _llm([
        {"role": "system", "content":
         "You are a QA Review Agent. Return only the numbers to keep."},
        {"role": "user", "content": _build_user_content(prompt, images)},
    ], temperature=0.1)

    keep = {int(n) for n in re.findall(r"\d+", response) if 1 <= int(n) <= len(tcs)}
    if not keep:
        _log("  ⚠️ Review returned no valid indices — keeping all.")
        return tcs
    _log(f"  ✔ Kept {len(keep)} · removed {len(tcs) - len(keep)}.")
    return [tc for i, tc in enumerate(tcs, 1) if i in keep]


# ── Parser ────────────────────────────────────────────────────────────────────
def _parse_tcs(content: str, story: dict) -> list:
    user_story = story["title"]
    ac_text    = story["acceptance_criteria"]
    out = []
    for block in content.split("---"):
        if "Title:" not in block:
            continue
        lines = block.strip().splitlines()
        get   = lambda k: next((l.split(k, 1)[1].strip() for l in lines if l.startswith(k)), None)

        # Priority
        try:
            p = int(get("Priority:") or 2)
            priority = p if p in (1, 2, 3, 4) else 2
        except (ValueError, TypeError):
            priority = 2

        # Steps (parse fenced code block)
        raw, in_block = [], False
        for line in lines:
            if line.strip().startswith("```"):
                in_block = not in_block
                continue
            if in_block:
                raw.append(line)
        steps_str = "\n".join(raw).strip()
        formatted = []
        try:
            if steps_str and not steps_str.startswith("["):
                steps_str = f"[{steps_str}]"
            for s in (ast.literal_eval(steps_str) if steps_str else []):
                if isinstance(s, dict) and "action" in s and "expected" in s:
                    formatted.append(f"{s['action']} -> {s['expected']}")
        except Exception:
            formatted = [steps_str] if steps_str else []

        out.append({
            "User Story":          user_story,
            "Acceptance Criteria": ac_text,
            "Test Type":           get("Test Type:") or "Functional",
            "Title":               get("Title:") or "Test Case",
            "Priority":            priority,
            "Steps":               "\n".join(formatted),
            "Status":              "Not Executed",
            "Comments":            "",
        })
    return out


# ── Save ──────────────────────────────────────────────────────────────────────
def _save(tcs: list) -> str:
    path = _OUTPUT_DIR / f"{datetime.now():%Y%m%d_%H%M%S}_generated_tcs.xlsx"
    df = pd.DataFrame(tcs)
    df.insert(0, "S.No.", range(1, len(df) + 1))
    df.to_excel(path, index=False)
    _log(f"  ✔ Saved {len(df)} test case(s) → {path.name}")
    return str(path)


# ── Pipelines ─────────────────────────────────────────────────────────────────
def _pipeline_standard(story: dict, suite_tcs: str = "") -> str:
    _log("🧠 Agent 1 · synthesizing test cases…")
    tcs = _parse_tcs(_agent_generate("agent1_generate.txt", story, suite_tcs), story)
    _log(f"  ✔ {len(tcs)} test case(s) generated.")
    _log("🔍 Agent 2 · reviewing for redundancy…")
    return _save(_agent_review(tcs, story))


def _pipeline_clone(story: dict, suite_tcs: str = "") -> str:
    _log("🧠 Cloning agent · adapting source suite…")
    tcs = _parse_tcs(_agent_generate("agent1_clone.txt", story, suite_tcs), story)
    _log(f"  ✔ {len(tcs)} test case(s) cloned.")
    return _save(tcs)


# ── ADO Test Manager ──────────────────────────────────────────────────────────
class ADOTestManager:
    """Thin wrapper over ADO Test Plans API. Lazily creates plan/suite as needed."""

    def __init__(self, org: str, project: str, pat: str, plan_name: str):
        self.org, self.project = org, project
        self.base    = f"https://dev.azure.com/{org}/{project}/_apis"
        self.auth    = HTTPBasicAuth("", pat)
        self.h       = {"Content-Type": "application/json"}
        self._suites: dict[str, int] = {}
        self.plan_id = self._ensure_plan(plan_name)

    def _get(self, path: str) -> dict:
        return requests.get(f"{self.base}/{path}", headers=self.h, auth=self.auth).json()

    def _post(self, path: str, payload: dict, ct: str = "application/json") -> dict:
        return requests.post(f"{self.base}/{path}", headers={"Content-Type": ct},
                             auth=self.auth, json=payload).json()

    def _ensure_plan(self, name: str) -> int:
        plans = self._get("testplan/plans?api-version=7.0").get("value", [])
        pid = next((p["id"] for p in plans if p["name"] == name), None)
        if pid:
            return pid
        return self._post("testplan/plans?api-version=7.0",
                          {"name": name, "areaPath": self.project, "iteration": self.project})["id"]

    def _ensure_suite(self, name: str) -> int:
        if name in self._suites:
            return self._suites[name]
        suites = self._get(f"testplan/plans/{self.plan_id}/suites?api-version=7.0").get("value", [])
        root = next((s for s in suites if s.get("suiteType") == "staticTestSuite"
                     and not s.get("parentSuite")), suites[0])
        sid  = next((s["id"] for s in suites if s["name"] == name), None)
        if not sid:
            sid = self._post(
                f"testplan/plans/{self.plan_id}/suites?api-version=7.0",
                {"suiteType": "staticTestSuite", "name": name,
                 "parentSuite": {"id": root["id"]}})["id"]
        self._suites[name] = sid
        return sid

    @staticmethod
    def _steps_xml(steps: list) -> str:
        body = "".join(
            f'<step id="{i}" type="ActionStep">'
            f'<parameterizedString isformatted="true">&lt;P&gt;{s["action"]}&lt;/P&gt;</parameterizedString>'
            f'<parameterizedString isformatted="true">&lt;P&gt;{s["expected"]}&lt;/P&gt;</parameterizedString>'
            f'<description/></step>'
            for i, s in enumerate(steps, 1)
        )
        return f'<steps id="0" last="{len(steps)}">{body}</steps>'

    def create_test_case(self, suite_name: str, title: str, steps: list,
                         priority: int = 2, story_id: str | None = None) -> int:
        sid   = self._ensure_suite(suite_name)
        patch = [
            {"op": "add", "path": "/fields/System.Title",                   "value": title},
            {"op": "add", "path": "/fields/Microsoft.VSTS.TCM.Steps",       "value": self._steps_xml(steps)},
            {"op": "add", "path": "/fields/Microsoft.VSTS.Common.Priority", "value": priority},
        ]
        if story_id:
            patch.append({"op": "add", "path": "/relations/-", "value": {
                "rel": "System.LinkTypes.Hierarchy-Reverse",
                "url": f"https://dev.azure.com/{self.org}/{self.project}/_apis/wit/workitems/{story_id}",
                "attributes": {"comment": "Auto-linked to parent story by TCS Synthesizer"},
            }})
        tc_id = self._post("wit/workitems/$Test Case?api-version=7.0",
                           patch, ct="application/json-patch+json")["id"]
        # Bind TC into suite
        requests.post(
            f"{self.base.replace('_apis', '_apis/test')}/Plans/{self.plan_id}"
            f"/Suites/{sid}/testcases/{tc_id}?api-version=5.0",
            headers=self.h, auth=self.auth,
        )
        suffix = f" → linked to story #{story_id}" if story_id else ""
        _log(f"    ✔ TC #{tc_id} · {title}{suffix}")
        return tc_id


# ── Upload ────────────────────────────────────────────────────────────────────
def _parse_steps_from_excel(text: Any) -> list:
    return [
        {"action": p[0].strip(), "expected": p[1].strip()}
        for line in str(text or "").splitlines() if "->" in line
        for p in [line.split("->", 1)]
    ]


def _upload(output_file: str, org: str, project: str, pat: str,
            plan_name: str, suite_name: str, story_id: str) -> tuple[int, int]:
    mgr = ADOTestManager(org, project, pat, plan_name)
    df  = pd.read_excel(output_file)
    uploaded = failed = 0
    for _, row in df.iterrows():
        if row.get("Status") == "Error":
            continue
        steps = _parse_steps_from_excel(row.get("Steps"))
        if not steps:
            failed += 1
            continue
        try:
            mgr.create_test_case(
                suite_name,
                row.get("Title", "Test Case"),
                steps,
                int(row.get("Priority", 2)),
                story_id=story_id,
            )
            uploaded += 1
        except Exception as e:
            failed += 1
            _log(f"    ❌ Upload failed · {row.get('Title', '?')} — {e}")
    return uploaded, failed


# ── Orchestrator ──────────────────────────────────────────────────────────────
def _run(label: str, story_id: str, org: str, project: str, pat: str,
         plan_override: str | None, suite_override: str | None,
         use_suite: tuple[str, str] | None = None,
         pipeline_fn: Callable = _pipeline_standard) -> dict:
    _log(f"⚡ {label} · Story #{story_id}")
    _log("🔗 Connecting to Azure DevOps…")
    story = _fetch_story(story_id, org, project, pat)
    title = story["title"] or f"Story {story_id}"
    _log(f"  ✔ Story fetched · \"{title}\"")

    suite_tcs = ""
    if use_suite:
        suite_tcs = _fetch_suite_tcs(org, project, pat, *use_suite)

    plan_name, suite_name = _build_plan_suite(story_id, title, plan_override, suite_override)
    _log(f"📋 Plan  · {plan_name}")
    _log(f"📁 Suite · {suite_name}")

    output_file = pipeline_fn(story, suite_tcs)
    generated   = len(pd.read_excel(output_file))

    _log(f"☁️ Uploading {generated} test case(s) to ADO…")
    uploaded, failed = _upload(output_file, org, project, pat, plan_name, suite_name, story_id)

    _hr()
    _log(f"✅ {label} complete · {uploaded} uploaded · {failed} failed")

    return {
        "story_id":        story_id,
        "story_title":     title,
        "plan_name":       plan_name,
        "suite_name":      suite_name,
        "generated_count": generated,
        "uploaded_count":  uploaded,
        "failed_count":    failed,
        "filename":        os.path.basename(output_file),
    }


# ── Public API ────────────────────────────────────────────────────────────────
def agentic_flow(story_id: str, org: str, project: str, pat: str,
                 plan_name_override: str | None = None,
                 suite_name_override: str | None = None) -> dict:
    return _run("Agentic Flow", story_id, org, project, pat,
                plan_name_override, suite_name_override)


def agentic_reference(story_id: str, org: str, project: str, pat: str,
                      ref_plan_id: str, ref_suite_id: str,
                      plan_name_override: str | None = None,
                      suite_name_override: str | None = None) -> dict:
    result = _run("Agentic Reference", story_id, org, project, pat,
                  plan_name_override, suite_name_override,
                  use_suite=(ref_plan_id, ref_suite_id))
    return {**result, "ref_plan_id": ref_plan_id, "ref_suite_id": ref_suite_id}


def agentic_clone(story_id: str, org: str, project: str, pat: str,
                  ref_plan_id: str, ref_suite_id: str,
                  plan_name_override: str | None = None,
                  suite_name_override: str | None = None) -> dict:
    result = _run("Agentic Clone", story_id, org, project, pat,
                  plan_name_override, suite_name_override,
                  use_suite=(ref_plan_id, ref_suite_id),
                  pipeline_fn=_pipeline_clone)
    return {**result, "ref_plan_id": ref_plan_id, "ref_suite_id": ref_suite_id}