import queue
import re, os, ast, base64
from html import unescape
from datetime import datetime
from pathlib import Path

import openai
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv

load_dotenv()
openai.api_type    = "azure"
openai.api_base    = os.getenv("AZURE_OPENAI_ENDPOINT")
openai.api_version = "2024-02-15-preview"
openai.api_key     = os.getenv("AZURE_OPENAI_KEY")

_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")

_REQUIRED_ENV = {
    "AZURE_OPENAI_ENDPOINT":   openai.api_base,
    "AZURE_OPENAI_KEY":        openai.api_key,
    "AZURE_OPENAI_DEPLOYMENT": _DEPLOYMENT,
}
_missing = [k for k, v in _REQUIRED_ENV.items() if not v]
if _missing:
    raise EnvironmentError(
        f"Missing required environment variable(s): {', '.join(_missing)}\n"
        f"Check your .env file is present and contains these keys."
    )

_PROMPTS_DIR = Path(__file__).parent / "prompts"

# ── Live log streaming ────────────────────────────────────────────────────────

_log_queue: "queue.Queue | None" = None

def set_log_queue(q: "queue.Queue | None") -> None:
    global _log_queue
    _log_queue = q

def _log(msg: str) -> None:
    print(msg)
    if _log_queue is not None:
        _log_queue.put(str(msg))


# ── Helpers ───────────────────────────────────────────────────────────────────

def _load_prompt(filename: str) -> str:
    return (_PROMPTS_DIR / filename).read_text(encoding="utf-8")

def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", unescape(str(text))).strip()

def _llm(messages: list, temperature: float = 0.3) -> str:
    r = openai.ChatCompletion.create(
        engine=_DEPLOYMENT, messages=messages, temperature=temperature,
    )
    return r.choices[0].message.content.strip()

def _save_output(tcs: list) -> str:
    os.makedirs("output", exist_ok=True)
    path = f"output/{datetime.now().strftime('%Y%m%d_%H%M%S')}_generated_tcs.xlsx"
    df = pd.DataFrame(tcs)
    df.insert(0, "S.No.", range(1, len(df) + 1))
    df.to_excel(path, index=False)
    _log(f"✅ Saved {len(df)} test cases → {path}")
    return path

def _build_plan_suite(story_id: str, title: str,
                      plan_name_override: str, suite_name_override: str) -> tuple:
    safe = re.sub(r"\s+", " ", re.sub(r"[^\w\s\-]", "", title).strip())[:50].strip()
    plan_name  = (plan_name_override  or "").strip() or f"TCS - #{story_id} - {safe}"
    suite_name = (suite_name_override or "").strip() or "TCS"
    return plan_name, suite_name


# ── ADO Fetchers ──────────────────────────────────────────────────────────────

_TEXT_EXTS = {".txt", ".md", ".markdown", ".json", ".xml", ".yaml", ".yml",
              ".csv", ".log", ".html", ".htm", ".rst", ".ini", ".toml"}
_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp"}
_IMAGE_MIME = {
    ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".png": "image/png",  ".gif": "image/gif", ".webp": "image/webp",
}


def _fetch_ado_comments(story_id, org, project, pat) -> str:
    url = (f"https://dev.azure.com/{org}/{project}/_apis/wit/workitems"
           f"/{story_id}/comments?api-version=7.1-preview.3")
    try:
        r = requests.get(url, headers={"Content-Type": "application/json"},
                         auth=HTTPBasicAuth("", pat), timeout=15)
        if r.status_code != 200:
            _log(f"  ⚠️ Could not fetch comments (HTTP {r.status_code}) — skipping.")
            return ""
        parts = []
        for c in r.json().get("comments", []):
            author = c.get("createdBy", {}).get("displayName", "Unknown")
            date   = (c.get("createdDate") or "")[:10]
            text   = _strip_html(c.get("text") or "")
            if text:
                parts.append(f"[{date}] {author}: {text}")
        if parts:
            _log(f"  💬 Fetched {len(parts)} comment(s).")
        return "\n".join(parts)
    except Exception as e:
        _log(f"  ⚠️ Error fetching comments: {e} — skipping.")
        return ""


def _fetch_ado_attachments(relations: list, pat: str) -> tuple:
    attachment_relations = [r for r in (relations or []) if r.get("rel") == "AttachedFile"]
    if not attachment_relations:
        return "", []
    _log(f"  📎 Found {len(attachment_relations)} attachment(s) — fetching content...")
    text_parts, images = [], []
    for rel in attachment_relations:
        url      = rel.get("url", "")
        filename = rel.get("attributes", {}).get("name", url.split("/")[-1])
        ext      = Path(filename).suffix.lower()
        try:
            resp = requests.get(url, auth=HTTPBasicAuth("", pat), timeout=60)
            if resp.status_code != 200:
                _log(f"    ⚠️ Could not download '{filename}' (HTTP {resp.status_code}) — skipping.")
                text_parts.append(f"[Attachment: {filename} — could not download]")
                continue
            if ext in _TEXT_EXTS:
                text = resp.content.decode("utf-8", errors="replace").strip()
                text_parts.append(f"[Attachment: {filename}]\n{text}")
                _log(f"    ✔ Text attachment: {filename} ({len(text)} chars)")
            elif ext in _IMAGE_EXTS:
                images.append({"media_type": _IMAGE_MIME[ext],
                                "data": base64.b64encode(resp.content).decode(),
                                "name": filename})
                _log(f"    🖼️ Image attachment: {filename}")
            else:
                text_parts.append(f"[Attachment: {filename} — binary file, content not included]")
                _log(f"    ℹ Binary attachment noted: {filename}")
        except Exception as e:
            _log(f"    ⚠️ Error fetching '{filename}': {e} — skipping.")
            text_parts.append(f"[Attachment: {filename} — fetch error]")
    return "\n\n".join(text_parts), images


def _extract_inline_images(html: str, pat: str) -> list:
    images = []
    for url in re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', html):
        if "dev.azure.com" not in url:
            continue
        try:
            resp = requests.get(url, auth=HTTPBasicAuth("", pat), timeout=30)
            if resp.status_code != 200:
                continue
            ct  = resp.headers.get("content-type", "image/png").split(";")[0].strip()
            ext = "." + ct.split("/")[-1].replace("jpeg", "jpg")
            if ext not in _IMAGE_EXTS:
                continue
            images.append({"media_type": ct,
                            "data": base64.b64encode(resp.content).decode(),
                            "name": url.split("/")[-1]})
            _log(f"    🖼️ Inline image fetched from story HTML.")
        except Exception:
            pass
    return images


def _fetch_ado_story(story_id, org, project, pat) -> dict:
    url = (f"https://dev.azure.com/{org}/{project}/_apis/wit/workitems"
           f"/{story_id}?$expand=all&api-version=7.0")
    r = requests.get(url, headers={"Content-Type": "application/json"},
                     auth=HTTPBasicAuth("", pat))
    if r.status_code != 200:
        raise Exception(f"ADO API error {r.status_code}: {r.text}")
    body = r.json()
    f    = body.get("fields", {})

    raw_desc = f.get("System.Description", "") or ""
    raw_ac   = f.get("Microsoft.VSTS.Common.AcceptanceCriteria", "") or ""

    inline_images                   = (_extract_inline_images(raw_desc, pat) +
                                       _extract_inline_images(raw_ac, pat))
    attachments_text, attach_images = _fetch_ado_attachments(body.get("relations") or [], pat)
    all_images = inline_images + attach_images
    if all_images:
        _log(f"  🖼️ {len(all_images)} image(s) collected total.")

    return {
        "id":                  story_id,
        "title":               f.get("System.Title", "").strip(),
        "description":         _strip_html(raw_desc),
        "acceptance_criteria": _strip_html(raw_ac),
        "comments":            _fetch_ado_comments(story_id, org, project, pat),
        "attachments_text":    attachments_text,
        "images":              all_images,
    }


def _fetch_suite_test_cases(org: str, project: str, pat: str,
                             plan_id: str, suite_id: str) -> str:
    """Fetch all TCs from an ADO suite and return as a formatted string."""
    _log(f"  📂 Fetching TCs from plan #{plan_id}, suite #{suite_id}...")
    auth = HTTPBasicAuth("", pat)
    h    = {"Content-Type": "application/json"}

    list_url = (f"https://dev.azure.com/{org}/{project}/_apis/testplan/Plans"
                f"/{plan_id}/Suites/{suite_id}/TestCase?api-version=7.0")
    r = requests.get(list_url, headers=h, auth=auth)
    if r.status_code != 200:
        raise Exception(f"Could not fetch suite TCs (HTTP {r.status_code}): {r.text}")

    refs = r.json().get("value", [])
    if not refs:
        _log("  ⚠️ Suite is empty — no TCs found.")
        return ""

    ids = [str(ref["workItem"]["id"]) for ref in refs]
    _log(f"  📋 Found {len(ids)} test case(s).")

    wi_url = (f"https://dev.azure.com/{org}/{project}/_apis/wit/workitems"
              f"?ids={','.join(ids)}"
              f"&fields=System.Title,Microsoft.VSTS.TCM.Steps&api-version=7.0")
    wi_r = requests.get(wi_url, headers=h, auth=auth)
    if wi_r.status_code != 200:
        raise Exception(f"Could not fetch TC details (HTTP {wi_r.status_code}): {wi_r.text}")

    parts = []
    for wi in wi_r.json().get("value", []):
        tc_title  = wi["fields"].get("System.Title", "Untitled")
        steps_xml = wi["fields"].get("Microsoft.VSTS.TCM.Steps", "") or ""
        raw       = re.findall(r'<parameterizedString[^>]*>(.*?)</parameterizedString>',
                               steps_xml, re.DOTALL)
        step_lines = [
            f"  Action: {_strip_html(raw[i]).strip()}\n  Expected: {_strip_html(raw[i+1]).strip()}"
            for i in range(0, len(raw) - 1, 2)
            if _strip_html(raw[i]).strip()
        ]
        parts.append(f"Title: {tc_title}\n" +
                     ("Steps:\n" + "\n".join(step_lines) if step_lines else "Steps: (none)"))

    _log(f"  ✅ TCs fetched and parsed.")
    return "\n\n---\n\n".join(parts)


# ── Agents ────────────────────────────────────────────────────────────────────

def _call_agent1(prompt_file: str, user_story: str, ac_text: str,
                 extra_context: str = "", comments: str = "",
                 attachments_text: str = "", suite_test_cases: str = "",
                 images: list = None) -> str:
    prompt = _load_prompt(prompt_file).format(
        user_story=user_story,
        acceptance_criteria=ac_text,
        extra_context=extra_context,
        comments=comments,
        attachments_text=attachments_text,
        reference_test_cases=suite_test_cases,
        clone_test_cases=suite_test_cases,
        figma_instruction="",
    )
    if images:
        _log(f"  🖼️ Sending {len(images)} image(s) to Agent 1.")
        content = [{"type": "text", "text": prompt}] + [
            {"type": "image_url",
             "image_url": {"url": f"data:{img['media_type']};base64,{img['data']}"}}
            for img in images
        ]
    else:
        content = prompt
    return _llm([
        {"role": "system", "content": "You are a QA engineer. Generate test cases in the EXACT format requested."},
        {"role": "user",   "content": content},
    ])


def _agent2_review(tcs: list, user_story: str = "", ac_text: str = "",
                   images: list = None) -> list:
    if not tcs:
        return tcs
    numbered = "\n".join(
        f"{i}. Title: {tc['Title']}\n   Steps:\n{tc['Steps']}"
        for i, tc in enumerate(tcs, 1)
    )
    prompt = _load_prompt("agent2_review.txt").format(
        test_cases=numbered,
        user_story=user_story,
        acceptance_criteria=ac_text,
    )
    if images:
        _log(f"  🖼️ Sending {len(images)} image(s) to Agent 2.")
        content = [{"type": "text", "text": prompt}] + [
            {"type": "image_url",
             "image_url": {"url": f"data:{img['media_type']};base64,{img['data']}"}}
            for img in images
        ]
    else:
        content = prompt
    response = _llm([
        {"role": "system", "content": "You are a QA Review Agent. Return only the numbers to keep."},
        {"role": "user",   "content": content},
    ], temperature=0.1)
    keep = {int(n) for n in re.findall(r"\d+", response) if 1 <= int(n) <= len(tcs)}
    if not keep:
        _log("⚠️ Review returned no valid indices — keeping all.")
        return tcs
    _log(f"✅ Review: kept {len(keep)}, removed {len(tcs) - len(keep)}.")
    return [tc for i, tc in enumerate(tcs, 1) if i in keep]


# ── Parser ────────────────────────────────────────────────────────────────────

def _parse_tcs(content: str, user_story: str, ac_text: str) -> list:
    result = []
    for block in content.split("---"):
        if "Title:" not in block:
            continue
        lines = block.strip().splitlines()
        get = lambda key: next((l.split(key, 1)[1].strip() for l in lines if l.startswith(key)), None)

        priority = 2
        try:
            p = int(get("Priority:") or 2)
            priority = p if p in (1, 2, 3, 4) else 2
        except (ValueError, TypeError):
            pass

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
            if not steps_str.startswith("["):
                steps_str = f"[{steps_str}]"
            for s in ast.literal_eval(steps_str):
                if isinstance(s, dict) and "action" in s and "expected" in s:
                    formatted.append(f"{s['action']} -> {s['expected']}")
        except Exception:
            formatted = [steps_str] if steps_str else []

        result.append({
            "User Story":          user_story,
            "Acceptance Criteria": ac_text,
            "Test Type":           get("Test Type:") or "Functional",
            "Title":               get("Title:") or "Test Case",
            "Priority":            priority,
            "Steps":               "\n".join(formatted),
            "Status":              "Not Executed",
            "Comments":            "",
        })
    return result


# ── Pipelines ─────────────────────────────────────────────────────────────────

def _run_pipeline(user_story: str, ac_text: str, extra_context: str = "",
                  comments: str = "", attachments_text: str = "",
                  suite_test_cases: str = "", images: list = None) -> str:
    _log("🤖 Agent 1: Generating test cases...")
    tcs = _parse_tcs(
        _call_agent1("agent1_generate.txt", user_story, ac_text, extra_context,
                     comments, attachments_text, suite_test_cases, images),
        user_story, ac_text,
    )
    _log(f"📋 Parsed {len(tcs)} test cases.")
    _log("🔍 Agent 2: Reviewing for redundancy...")
    return _save_output(_agent2_review(tcs, user_story, ac_text, images))


def _run_clone_pipeline(user_story: str, ac_text: str, extra_context: str = "",
                        comments: str = "", attachments_text: str = "",
                        suite_test_cases: str = "", images: list = None) -> str:
    _log("🤖 Agent 1 (clone): Generating test cases...")
    tcs = _parse_tcs(
        _call_agent1("agent1_clone.txt", user_story, ac_text, extra_context,
                     comments, attachments_text, suite_test_cases, images),
        user_story, ac_text,
    )
    _log(f"📋 Parsed {len(tcs)} test cases.")
    return _save_output(tcs)


# ── ADO Test Manager ──────────────────────────────────────────────────────────

class ADOTestManager:
    def __init__(self, org, proj, pat, plan_name):
        self.org, self.proj = org, proj
        self.base    = f"https://dev.azure.com/{org}/{proj}/_apis"
        self.auth    = HTTPBasicAuth("", pat)
        self.h       = {"Content-Type": "application/json"}
        self._suites = {}
        self.plan_id = self._setup_plan(plan_name)

    def _setup_plan(self, name: str) -> int:
        plans = requests.get(f"{self.base}/testplan/plans?api-version=7.0",
                             headers=self.h, auth=self.auth).json().get("value", [])
        pid = next((p["id"] for p in plans if p["name"] == name), None)
        if not pid:
            pid = requests.post(f"{self.base}/testplan/plans?api-version=7.0",
                                headers=self.h, auth=self.auth,
                                json={"name": name, "areaPath": self.proj,
                                      "iteration": self.proj}).json()["id"]
        return pid

    def _get_suite(self, name: str) -> int:
        if name in self._suites:
            return self._suites[name]
        suites = requests.get(f"{self.base}/testplan/plans/{self.plan_id}/suites?api-version=7.0",
                              headers=self.h, auth=self.auth).json().get("value", [])
        root = next((s for s in suites if s.get("suiteType") == "staticTestSuite"
                     and not s.get("parentSuite")), suites[0])
        sid = next((s["id"] for s in suites if s["name"] == name), None)
        if not sid:
            sid = requests.post(f"{self.base}/testplan/plans/{self.plan_id}/suites?api-version=7.0",
                                headers=self.h, auth=self.auth,
                                json={"suiteType": "staticTestSuite", "name": name,
                                      "parentSuite": {"id": root["id"]}}).json()["id"]
        self._suites[name] = sid
        return sid

    def create_test_case(self, suite_name: str, title: str, steps: list,
                         priority: int = 2, story_id: str = None) -> int:
        sid = self._get_suite(suite_name)
        xml = f'<steps id="0" last="{len(steps)}">' + "".join(
            f'<step id="{i}" type="ActionStep">'
            f'<parameterizedString isformatted="true">&lt;P&gt;{s["action"]}&lt;/P&gt;</parameterizedString>'
            f'<parameterizedString isformatted="true">&lt;P&gt;{s["expected"]}&lt;/P&gt;</parameterizedString>'
            f'<description/></step>'
            for i, s in enumerate(steps, 1)
        ) + "</steps>"
        patch = [
            {"op": "add", "path": "/fields/System.Title",                   "value": title},
            {"op": "add", "path": "/fields/Microsoft.VSTS.TCM.Steps",       "value": xml},
            {"op": "add", "path": "/fields/Microsoft.VSTS.Common.Priority", "value": priority},
        ]
        if story_id:
            patch.append({
                "op": "add", "path": "/relations/-",
                "value": {
                    "rel": "System.LinkTypes.Hierarchy-Reverse",
                    "url": f"https://dev.azure.com/{self.org}/{self.proj}/_apis/wit/workitems/{story_id}",
                    "attributes": {"comment": "Auto-linked to parent user story by TCS Synthesizer"},
                },
            })
        tc_id = requests.post(
            f"{self.base}/wit/workitems/$Test Case?api-version=7.0",
            headers={"Content-Type": "application/json-patch+json"},
            auth=self.auth, json=patch,
        ).json()["id"]
        requests.post(
            f"https://dev.azure.com/{self.org}/{self.proj}/_apis/test"
            f"/Plans/{self.plan_id}/Suites/{sid}/testcases/{tc_id}?api-version=5.0",
            headers=self.h, auth=self.auth,
        )
        _log(f"  ✔ TC #{tc_id}: '{title}'" + (f" → linked to story #{story_id}" if story_id else ""))
        return tc_id


# ── Public API ────────────────────────────────────────────────────────────────

def _upload_to_ado(output_file: str, org: str, project: str, pat: str,
                   plan_name: str, suite_name: str, story_id: str) -> tuple:
    mgr = ADOTestManager(org, project, pat, plan_name)
    df  = pd.read_excel(output_file)
    uploaded, failed = 0, 0
    for _, row in df.iterrows():
        if row.get("Status") == "Error":
            continue
        steps = [
            {"action": p[0].strip(), "expected": p[1].strip()}
            for line in str(row.get("Steps", "")).strip().splitlines()
            if "->" in line
            for p in [line.split("->", 1)]
        ]
        if not steps:
            failed += 1
            continue
        try:
            mgr.create_test_case(suite_name, row.get("Title", "Test Case"),
                                 steps, int(row.get("Priority", 2)), story_id=story_id)
            uploaded += 1
        except Exception as e:
            failed += 1
            _log(f"  ❌ Failed to upload: {row.get('Title', '?')} — {e}")
    _log(f"☁️ Upload complete — {uploaded} uploaded, {failed} failed.")
    return uploaded, failed


def _run_agentic_core(label: str, story_id: str, org: str, project: str, pat: str,
                      plan_name_override: str, suite_name_override: str,
                      s: dict, suite_tcs: str = "", pipeline_fn=None) -> dict:
    if pipeline_fn is None:
        pipeline_fn = _run_pipeline
    title   = s["title"] or f"Story {story_id}"
    ac_text = s["acceptance_criteria"] or "No acceptance criteria provided."
    extra   = f"Description: {s['description']}" if s["description"] else ""

    plan_name, suite_name = _build_plan_suite(story_id, title, plan_name_override, suite_name_override)
    _log(f"  📋 Test Plan : '{plan_name}'")
    _log(f"  📁 Test Suite: '{suite_name}'")
    _log(f"  🧠 Generating test cases...")

    output_file = pipeline_fn(title, ac_text, extra,
                              s["comments"], s["attachments_text"],
                              suite_tcs, s["images"])
    generated = len(pd.read_excel(output_file))

    _log(f"  ☁️ Uploading {generated} test cases to ADO...")
    uploaded, failed = _upload_to_ado(output_file, org, project, pat,
                                      plan_name, suite_name, story_id)
    _log(f"✅ {label} complete — {uploaded} uploaded, {failed} failed.")

    return {
        "story_id":        story_id, "story_title":    title,
        "plan_name":       plan_name, "suite_name":    suite_name,
        "generated_count": generated, "uploaded_count": uploaded,
        "failed_count":    failed,   "filename":       os.path.basename(output_file),
    }


def agentic_flow(story_id: str, org: str, project: str, pat: str,
                 plan_name_override: str = None, suite_name_override: str = None) -> dict:
    _log(f"⚡ Agentic Flow started for story #{story_id}")
    s = _fetch_ado_story(story_id, org, project, pat)
    _log(f"  ✅ Story fetched: '{s['title']}'")
    return _run_agentic_core("Agentic Flow", story_id, org, project, pat,
                             plan_name_override, suite_name_override, s)


def agentic_reference(story_id: str, org: str, project: str, pat: str,
                      ref_plan_id: str, ref_suite_id: str,
                      plan_name_override: str = None, suite_name_override: str = None) -> dict:
    _log(f"⚡ Agentic Reference started for story #{story_id}")
    s = _fetch_ado_story(story_id, org, project, pat)
    _log(f"  ✅ Story fetched: '{s['title']}'")
    suite_tcs = _fetch_suite_test_cases(org, project, pat, ref_plan_id, ref_suite_id)
    result = _run_agentic_core("Agentic Reference", story_id, org, project, pat,
                               plan_name_override, suite_name_override, s, suite_tcs)
    return {**result, "ref_plan_id": ref_plan_id, "ref_suite_id": ref_suite_id}


def agentic_clone(story_id: str, org: str, project: str, pat: str,
                  ref_plan_id: str, ref_suite_id: str,
                  plan_name_override: str = None, suite_name_override: str = None) -> dict:
    _log(f"⚡ Agentic Clone started for story #{story_id}")
    s = _fetch_ado_story(story_id, org, project, pat)
    _log(f"  ✅ Story fetched: '{s['title']}'")
    suite_tcs = _fetch_suite_test_cases(org, project, pat, ref_plan_id, ref_suite_id)
    result = _run_agentic_core("Agentic Clone", story_id, org, project, pat,
                               plan_name_override, suite_name_override, s,
                               suite_tcs, pipeline_fn=_run_clone_pipeline)
    return {**result, "ref_plan_id": ref_plan_id, "ref_suite_id": ref_suite_id}