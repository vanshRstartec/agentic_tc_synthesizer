import re, os, ast, base64
from datetime import datetime
from pathlib import Path

import openai
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv

load_dotenv()
openai.api_type = "azure"
openai.api_base = os.getenv("AZURE_OPENAI_ENDPOINT")
openai.api_version = "2024-02-15-preview"
openai.api_key = os.getenv("AZURE_OPENAI_KEY")

_PROMPTS_DIR = Path(__file__).parent / "prompts"


# ── Helpers ──────────────────────────────────────────────────────────────────

def _load_prompt(filename: str) -> str:
    return (_PROMPTS_DIR / filename).read_text(encoding="utf-8")

def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", str(text)).strip()

def _load_image_b64(image_path: str) -> str:
    if not image_path:
        return None
    try:
        return base64.b64encode(Path(image_path).read_bytes()).decode()
    except Exception as e:
        print(f"⚠️ Could not load image {image_path}: {e}")
        return None

def _llm(messages: list, temperature: float = 0.3) -> str:
    r = openai.ChatCompletion.create(
        engine=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
        messages=messages,
        temperature=temperature,
    )
    return r.choices[0].message.content.strip()

def _save_output(tcs: list, path: str = None) -> str:
    os.makedirs("output", exist_ok=True)
    path = path or f"output/{datetime.now().strftime('%Y%m%d_%H%M%S')}_generated_tcs.xlsx"
    df = pd.DataFrame(tcs)
    df.insert(0, "S.No.", range(1, len(df) + 1))
    df.to_excel(path, index=False)
    print(f"✅ Saved {len(df)} test cases → {path}")
    return path


# ── ADO Story Fetcher ────────────────────────────────────────────────────────

def _fetch_ado_story(story_id, org, project, pat) -> dict:
    url = f"https://dev.azure.com/{org}/{project}/_apis/wit/workitems/{story_id}?$expand=all&api-version=7.0"
    r = requests.get(url, headers={"Content-Type": "application/json"}, auth=HTTPBasicAuth("", pat))
    if r.status_code != 200:
        raise Exception(f"ADO API error {r.status_code}: {r.text}")
    f = r.json().get("fields", {})
    return {
        "id": story_id,
        "title": f.get("System.Title", "").strip(),
        "description": _strip_html(f.get("System.Description", "") or ""),
        "acceptance_criteria": _strip_html(f.get("Microsoft.VSTS.Common.AcceptanceCriteria", "") or ""),
    }


# ── Agent 1: Generate ────────────────────────────────────────────────────────

def _agent1_generate(user_story: str, ac_text: str, extra_context: str = "", image_b64: str = None) -> str:
    prompt = _load_prompt("agent1_generate.txt").format(
        user_story=user_story,
        acceptance_criteria=ac_text,
        extra_context=extra_context,
        figma_instruction=(
            "Use the provided Figma screenshot for UI reference. "
            "Include UI test cases to verify content and layout only."
        ) if image_b64 else "",
    )
    user_msg = (
        [{"type": "text", "text": prompt},
         {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}}]
        if image_b64 else prompt
    )
    return _llm([
        {"role": "system", "content": "You are a QA engineer. Generate test cases in the EXACT format requested."},
        {"role": "user", "content": user_msg},
    ])


# ── Parser ───────────────────────────────────────────────────────────────────

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
            "User Story": user_story,
            "Acceptance Criteria": ac_text,
            "Test Type": get("Test Type:") or "Functional",
            "Title": get("Title:") or "Test Case",
            "Priority": priority,
            "Steps": "\n".join(formatted),
            "Status": "Not Executed",
            "Comments": "",
        })
    return result


# ── Agent 2: Review ──────────────────────────────────────────────────────────

def _agent2_review(tcs: list) -> list:
    if not tcs:
        return tcs
    numbered = "\n".join(
        f"{i}. Title: {tc['Title']}\n   Steps:\n{tc['Steps']}"
        for i, tc in enumerate(tcs, 1)
    )
    response = _llm([
        {"role": "system", "content": "You are a QA Review Agent. Return only the numbers to keep."},
        {"role": "user", "content": _load_prompt("agent2_review.txt").format(test_cases=numbered)},
    ], temperature=0.1)

    keep = {int(n) for n in re.findall(r"\d+", response) if 1 <= int(n) <= len(tcs)}
    if not keep:
        print("⚠️ Review returned no valid indices — keeping all.")
        return tcs
    print(f"✅ Review: kept {len(keep)}, removed {len(tcs) - len(keep)}.")
    return [tc for i, tc in enumerate(tcs, 1) if i in keep]


# ── Core pipeline ────────────────────────────────────────────────────────────

def _run_pipeline(user_story: str, ac_text: str, extra_context: str = "", image_b64: str = None) -> str:
    print("🤖 Agent 1: Generating...")
    tcs = _parse_tcs(_agent1_generate(user_story, ac_text, extra_context, image_b64), user_story, ac_text)
    print(f"📋 Parsed {len(tcs)} test cases.")
    print("🔍 Agent 2: Reviewing...")
    return _save_output(_agent2_review(tcs))


# ── Public API ───────────────────────────────────────────────────────────────

def generate_from_excel(input_file: str, output_file: str = None, image_path: str = None) -> str:
    df = pd.read_excel(input_file)
    image_b64 = _load_image_b64(image_path)

    groups = {}
    for _, row in df.iterrows():
        story = str(row.get("User Story", "") or "").strip() or "Unknown Story"
        ac = str(row.get("Acceptance Criteria", "") or "").strip()
        if pd.isna(row.get("User Story")) and pd.isna(row.get("Acceptance Criteria")):
            continue
        groups.setdefault(story, {"acs": [], "rows": []})
        groups[story]["acs"].append(ac)
        groups[story]["rows"].append(row)

    print(f"📊 {len(groups)} unique user stories.")
    all_tcs = []

    for i, (story, data) in enumerate(groups.items(), 1):
        print(f"\n  Story {i}/{len(groups)}: {story[:60]}...")
        first = data["rows"][0]
        extra = "\n".join(
            f"{c}: {first.get(c)}" for c in
            ["Feature/Module", "Priority", "Risk Level", "Preconditions",
             "Test Environment", "Generic Test Data", "Comments/Notes"]
            if pd.notna(first.get(c)) and str(first.get(c, "")).strip()
        )
        ac_text = "\n".join(f"AC {j}: {ac}" for j, ac in enumerate(data["acs"], 1))
        raw = _agent1_generate(story, ac_text, extra, image_b64)
        all_tcs.extend(_parse_tcs(raw, story, ac_text))

    print(f"\n📋 Total: {len(all_tcs)} | 🔍 Agent 2 reviewing...")
    all_tcs = _agent2_review(all_tcs)

    path = output_file or f"output/{datetime.now().strftime('%Y%m%d_%H%M%S')}_generated_tcs.xlsx"
    os.makedirs(os.path.dirname(path) or "output", exist_ok=True)
    df_out = pd.DataFrame(all_tcs)
    df_out.insert(0, "S.No.", range(1, len(df_out) + 1))
    df_out.to_excel(path, index=False)
    print(f"✅ {len(df_out)} test cases → {path}")
    return path


def generate_from_ado(story_id: str, org: str, project: str, pat: str, image_path: str = None) -> str:
    print(f"\n🔗 Fetching ADO story #{story_id}...")
    s = _fetch_ado_story(story_id, org, project, pat)
    title = s["title"] or f"User Story #{story_id}"
    ac_text = s["acceptance_criteria"] or "No acceptance criteria provided."
    extra = f"Description: {s['description']}" if s["description"] else ""
    print(f"  ✅ '{title}'")
    return _run_pipeline(title, ac_text, extra, _load_image_b64(image_path))


def upload_test_cases_ado(
    excel_file: str,
    org: str,
    proj: str,
    pat: str,
    plan_name: str,
    suite_name: str,
    story_id: str = None,
):
    df = pd.read_excel(excel_file)
    mgr = ADOTestManager(org, proj, pat, plan_name)
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
            print(f"❌ {row.get('Title', '?')}: {e}")
    print(f"✅ Uploaded {uploaded}/{len(df)} ({failed} failed)")
    return uploaded, failed


def agentic_flow(
    story_id: str,
    org: str,
    project: str,
    pat: str,
    plan_name_override: str = None,
    suite_name_override: str = None,
) -> dict:
    """
    Override logic:
    - plan_name_override provided  → use it as plan name
    - suite_name_override provided → use it as suite name
    - neither provided             → auto-generate plan name, suite = "TCS"
    - suite only provided          → auto-generate plan name, use override suite
    - plan only provided           → use override plan name, suite = "TCS"
    """
    print(f"\n⚡ Agentic Flow: story #{story_id}")
    s = _fetch_ado_story(story_id, org, project, pat)
    title = s["title"] or f"Story {story_id}"

    # Build auto plan name from story title (used when no override supplied)
    safe = re.sub(r"\s+", " ", re.sub(r"[^\w\s\-]", "", title).strip())[:60].strip()
    auto_plan_name = f"TCS - {safe}"

    plan_name = plan_name_override.strip() if plan_name_override and plan_name_override.strip() else auto_plan_name
    suite_name = suite_name_override.strip() if suite_name_override and suite_name_override.strip() else "TCS"

    print(f"  Plan: '{plan_name}' | Suite: '{suite_name}'")

    output_file = generate_from_ado(story_id, org, project, pat)
    generated = len(pd.read_excel(output_file))
    uploaded, failed = upload_test_cases_ado(output_file, org, project, pat, plan_name, suite_name, story_id=story_id)

    return {
        "story_id": story_id,
        "story_title": title,
        "plan_name": plan_name,
        "suite_name": suite_name,
        "generated_count": generated,
        "uploaded_count": uploaded,
        "failed_count": failed,
        "filename": os.path.basename(output_file),
    }


# ── Figma utility ────────────────────────────────────────────────────────────

def download_image(url: str, figma_token: str):
    m = re.search(r"figma\.com/(?:design|file)/([a-zA-Z0-9]+)/[^?]*\?node-id=([\d-]+)", url)
    if not m:
        print("❌ Failed to parse Figma URL")
        return None
    file_key, node_id = m.group(1), m.group(2).replace("-", ":")
    png_url = requests.get(
        f"https://api.figma.com/v1/images/{file_key}?ids={node_id}&format=png",
        headers={"X-Figma-Token": figma_token}
    ).json()["images"][node_id]
    Path("figma.png").write_bytes(requests.get(png_url).content)
    return "figma.png"


# ── ADO Test Manager ─────────────────────────────────────────────────────────

class ADOTestManager:
    def __init__(self, org, proj, pat, plan_name):
        self.org, self.proj = org, proj
        self.base = f"https://dev.azure.com/{org}/{proj}/_apis"
        self.auth = HTTPBasicAuth("", pat)
        self.h = {"Content-Type": "application/json"}
        self._suites: dict = {}
        self.plan_id = self._setup_plan(plan_name)

    def _setup_plan(self, name: str) -> int:
        plans = requests.get(f"{self.base}/testplan/plans?api-version=7.0",
                             headers=self.h, auth=self.auth).json().get("value", [])
        pid = next((p["id"] for p in plans if p["name"] == name), None)
        if not pid:
            pid = requests.post(f"{self.base}/testplan/plans?api-version=7.0",
                                headers=self.h, auth=self.auth,
                                json={"name": name, "areaPath": self.proj, "iteration": self.proj}).json()["id"]
        return pid

    def _get_suite(self, name: str) -> int:
        if name in self._suites:
            return self._suites[name]
        suites = requests.get(f"{self.base}/testplan/plans/{self.plan_id}/suites?api-version=7.0",
                              headers=self.h, auth=self.auth).json().get("value", [])
        root = next((s for s in suites if s.get("suiteType") == "staticTestSuite" and not s.get("parentSuite")),
                    suites[0])
        sid = next((s["id"] for s in suites if s["name"] == name), None)
        if not sid:
            sid = requests.post(f"{self.base}/testplan/plans/{self.plan_id}/suites?api-version=7.0",
                                headers=self.h, auth=self.auth,
                                json={"suiteType": "staticTestSuite", "name": name,
                                      "parentSuite": {"id": root["id"]}}).json()["id"]
        self._suites[name] = sid
        return sid

    def create_test_case(
        self,
        suite_name: str,
        title: str,
        steps: list,
        priority: int = 2,
        story_id: str = None,
    ) -> int:
        sid = self._get_suite(suite_name)
        xml = f'<steps id="0" last="{len(steps)}">' + "".join(
            f'<step id="{i}" type="ActionStep">'
            f'<parameterizedString isformatted="true">&lt;P&gt;{s["action"]}&lt;/P&gt;</parameterizedString>'
            f'<parameterizedString isformatted="true">&lt;P&gt;{s["expected"]}&lt;/P&gt;</parameterizedString>'
            f'<description/></step>'
            for i, s in enumerate(steps, 1)
        ) + "</steps>"

        patch = [
            {"op": "add", "path": "/fields/System.Title",                      "value": title},
            {"op": "add", "path": "/fields/Microsoft.VSTS.TCM.Steps",          "value": xml},
            {"op": "add", "path": "/fields/Microsoft.VSTS.Common.Priority",    "value": priority},
        ]

        # Link the test case to the parent user story as a child work item
        if story_id:
            story_url = (
                f"https://dev.azure.com/{self.org}/{self.proj}"
                f"/_apis/wit/workitems/{story_id}"
            )
            patch.append({
                "op": "add",
                "path": "/relations/-",
                "value": {
                    "rel": "System.LinkTypes.Hierarchy-Reverse",
                    "url": story_url,
                    "attributes": {"comment": "Auto-linked to parent user story by TCS Synthesizer"},
                },
            })

        tc_id = requests.post(
            f"{self.base}/wit/workitems/$Test Case?api-version=7.0",
            headers={"Content-Type": "application/json-patch+json"},
            auth=self.auth,
            json=patch,
        ).json()["id"]

        requests.post(
            f"https://dev.azure.com/{self.org}/{self.proj}/_apis/test"
            f"/Plans/{self.plan_id}/Suites/{sid}/testcases/{tc_id}?api-version=5.0",
            headers=self.h, auth=self.auth,
        )
        link_note = f" → linked to story #{story_id}" if story_id else ""
        print(f"  Created TC #{tc_id} in '{suite_name}'{link_note}")
        return tc_id