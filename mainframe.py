import re

import openai
import pandas as pd
import os
import ast
import requests
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv

load_dotenv()
openai.api_type = "azure"
openai.api_base = os.getenv("AZURE_OPENAI_ENDPOINT")
openai.api_version = "2024-02-15-preview"
openai.api_key = os.getenv("AZURE_OPENAI_KEY")


class ADOTestManager:
    def __init__(self, org, proj, pat, plan_name):
        self.org, self.proj, self.base = org, proj, f"https://dev.azure.com/{org}/{proj}/_apis"
        self.auth, self.h, self.suites = HTTPBasicAuth('', pat), {"Content-Type": "application/json"}, {}
        self.plan_id = self._setup_plan(plan_name)

    def _setup_plan(self, plan_name):
        r = requests.get(f"{self.base}/testplan/plans?api-version=7.0", headers=self.h, auth=self.auth)
        plan_id = next((p['id'] for p in r.json().get('value', []) if p['name'] == plan_name), None)
        if not plan_id:
            r = requests.post(f"{self.base}/testplan/plans?api-version=7.0", headers=self.h, auth=self.auth,
                              json={"name": plan_name, "areaPath": self.proj, "iteration": self.proj})
            plan_id = r.json()['id']
        return plan_id

    def _get_suite(self, suite_name):
        if suite_name in self.suites: return self.suites[suite_name]
        r = requests.get(f"{self.base}/testplan/plans/{self.plan_id}/suites?api-version=7.0", headers=self.h,
                         auth=self.auth)
        suites = r.json().get('value', [])
        root = next((s for s in suites if s.get('suiteType') == 'staticTestSuite' and s.get('parentSuite') is None),
                    suites[0])
        suite_id = next((s['id'] for s in suites if s['name'] == suite_name), None)
        if not suite_id:
            r = requests.post(f"{self.base}/testplan/plans/{self.plan_id}/suites?api-version=7.0", headers=self.h,
                              auth=self.auth,
                              json={"suiteType": "staticTestSuite", "name": suite_name,
                                    "parentSuite": {"id": root['id']}})
            suite_id = r.json()['id']
        self.suites[suite_name] = suite_id
        return suite_id

    def create_test_case(self, suite_name, title, steps, priority=2):
        suite_id = self._get_suite(suite_name)
        steps_xml = f'<steps id="0" last="{len(steps)}">'
        for i, s in enumerate(steps, 1):
            steps_xml += f'<step id="{i}" type="ActionStep"><parameterizedString isformatted="true">&lt;P&gt;{s["action"]}&lt;/P&gt;</parameterizedString><parameterizedString isformatted="true">&lt;P&gt;{s["expected"]}&lt;/P&gt;</parameterizedString><description/></step>'
        steps_xml += '</steps>'
        payload = [{"op": "add", "path": "/fields/System.Title", "value": title},
                   {"op": "add", "path": "/fields/Microsoft.VSTS.TCM.Steps", "value": steps_xml},
                   {"op": "add", "path": "/fields/Microsoft.VSTS.Common.Priority", "value": priority}]
        r = requests.post(f"{self.base}/wit/workitems/$Test Case?api-version=7.0",
                          headers={"Content-Type": "application/json-patch+json"}, auth=self.auth, json=payload)
        tc_id = r.json()['id']
        requests.post(
            f"https://dev.azure.com/{self.org}/{self.proj}/_apis/test/Plans/{self.plan_id}/Suites/{suite_id}/testcases/{tc_id}?api-version=5.0",
            headers=self.h, auth=self.auth)
        print(f"Created Test Case #{tc_id} in suite '{suite_name}'")
        return tc_id


def generate_test_cases(input_file, output_file=None, image_path=None):
    from datetime import datetime
    import base64
    if output_file is None:
        os.makedirs("output", exist_ok=True)
        output_file = f"output/{datetime.now().strftime('%Y%m%d_%H%M%S')}_generated_tcs.xlsx"

    df = pd.read_excel(input_file)
    all_generated_tcs = []

    if output_dir := os.path.dirname(output_file):
        os.makedirs(output_dir, exist_ok=True)

    figma_image_base64 = None
    if image_path:
        try:
            with open(image_path, "rb") as image_file:
                figma_image_base64 = base64.b64encode(image_file.read()).decode('utf-8')
            print(f"✅ Found {image_path} - will include in API calls")
        except Exception as e:
            print(f"⚠️ Error reading {image_path}: {e}")

    # Step 1: Group acceptance criteria by user story
    print(f"📊 Grouping acceptance criteria by user story...")
    story_groups = {}

    for idx, row in df.iterrows():
        user_story = row.get("User Story", "")
        ac = row.get("Acceptance Criteria", "")

        # Skip if both are empty
        if pd.isna(user_story) and pd.isna(ac):
            continue

        # Handle empty user story (use previous story or mark as unknown)
        if pd.isna(user_story) or str(user_story).strip() == "":
            user_story = "Unknown Story"

        user_story = str(user_story).strip()

        # Initialize story group if not exists
        if user_story not in story_groups:
            story_groups[user_story] = {
                "acceptance_criteria": [],
                "context_rows": []
            }

        # Add AC and context for this story
        story_groups[user_story]["acceptance_criteria"].append(str(ac) if pd.notna(ac) else "")
        story_groups[user_story]["context_rows"].append(row)

    print(f"✅ Found {len(story_groups)} unique user stories")

    # Step 2: Generate test cases for each user story (with all its ACs at once)
    print(f"\n🤖 Agent 1: Generating test cases story-wise...")

    for story_num, (user_story, story_data) in enumerate(story_groups.items(), 1):
        acs = story_data["acceptance_criteria"]
        context_rows = story_data["context_rows"]

        print(f"\n  Processing Story {story_num}/{len(story_groups)}: {user_story[:50]}...")
        print(f"  └─ {len(acs)} acceptance criteria to process")

        # Build context from the first row (assuming similar context for same story)
        first_row = context_rows[0]
        context = "\n".join([f"{col}: {first_row.get(col, '')}" for col in
                             ["Feature/Module", "Priority", "Risk Level", "Preconditions", "Test Environment",
                              "Generic Test Data", "Comments/Notes"] if
                             pd.notna(first_row.get(col, "")) and str(first_row.get(col, "")).strip()])

        # Build combined acceptance criteria text
        ac_text = ""
        for i, ac in enumerate(acs, 1):
            ac_text += f"\nAC {i}: {ac}"

        prompt_text = f"""You are a QA Engineer at Energy Utility Company. Your responsibility is to create test cases for assistance tools developed for Customer Service Representatives (CSRs). These tools are designed to help CSRs communicate effectively and efficiently with customers.
        
        You would have to generate test cases for the below requirement:

User Story: {user_story}

Acceptance Criteria:
{ac_text}

{context}

{"Use the provided Figma screenshot for UI reference, and make sure to include UI test cases to verify content and layout only." if figma_image_base64 else ""}

Instructions:
 - CREATE A MIX OF POSITIVE/NEGATIVE/EDGE CASES (if needed) BASED ON THE PROVIDED USER STORY AND ACCEPTANCE CRITERIA ONLY
 - DONT ASSUME THINGS THAT ARE NOT PROVIDED IN THE ACCEPTANCE CRITERIA. 
 - TEST CASES COUNT DOES NOT MATTER UNLESS THE STORY AND THE ACCEPTANCE CRITERIA IS MET. 
 - REFER QUALITY OVER QUANTITY.
 - DONT MAKE ANY SCENARIO WHICH WILL REQUIRE ADDITIONAL DEVELOPMENT AND IS NOT PRESENT IN THE ACCEPTANCE CRITERIA
**MAKE SURE YOU COVER ALL THE SCENARIOS MENTIONED IN THE ACCEPTANCE CRITERIA. IF SOMETHING IS NOT CLEAR, SKIP IT RATHER THAN MAKING ASSUMPTIONS.**

For EACH test case, use this EXACT format:

Test Type: [Positive/Negative/Edge Case(if required)]
Title: [Clear test case title]
Priority: [1-4]
Steps:
```
{{'action': '[step action]', 'expected': '[step expected]'}},
{{'action': '[step action]', 'expected': '[step expected]'}},
{{'action': '[step action]', 'expected': '[step expected]'}}
```
---

Each step must be a dictionary with 'action' and 'expected' keys.
Between each test case there should be a line with only --- to separate them.
"""

        try:
            messages = [
                {"role": "system",
                 "content": "You are a QA engineer. Generate test cases in the EXACT format requested."}
            ]

            if figma_image_base64:
                messages.append({
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt_text},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{figma_image_base64}"
                            }
                        }
                    ]
                })
            else:
                messages.append({"role": "user", "content": prompt_text})

            response = openai.ChatCompletion.create(
                engine=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                messages=messages,
                temperature=0.3,
                max_tokens=2500  # Increased for multiple ACs
            )
            content = response.choices[0].message.content.strip()

            # Store generated test cases with their context
            all_generated_tcs.append({
                "user_story": user_story,
                "acceptance_criteria": ac_text,  # Store all ACs
                "generated_content": content
            })
            print(f"  ✓ Generated TCs for story {story_num}/{len(story_groups)}")

        except Exception as e:
            print(f"  ❌ Error generating TCs for story {story_num}: {e}")
            all_generated_tcs.append({
                "user_story": user_story,
                "acceptance_criteria": ac_text,
                "generated_content": f"ERROR: {str(e)}"
            })

    # Step 3: Parse all generated test cases
    print(f"\n📋 Parsing {len(all_generated_tcs)} generated test case sets...")
    all_parsed_tcs = []

    for tc_set in all_generated_tcs:
        content = tc_set["generated_content"]
        for block in [b for b in content.split("---") if "Title:" in b]:
            lines = block.strip().split("\n")
            test_type = next((l.replace("Test Type:", "").strip() for l in lines if "Test Type:" in l), "Functional")
            title = next((l.replace("Title:", "").strip() for l in lines if "Title:" in l), "Test Case")
            priority = next((l.replace("Priority:", "").strip() for l in lines if "Priority:" in l), "2")
            try:
                priority = int(priority)
                if priority not in [1, 2, 3, 4]: priority = 2
            except:
                priority = 2
            steps, in_code_block = "", False
            for line in lines:
                if line.strip().startswith("```"): in_code_block = not in_code_block; continue
                if in_code_block: steps += line + "\n"
            steps_formatted = []
            try:
                if not steps.strip().startswith('['): steps = '[' + steps.strip() + ']'
                steps_list = ast.literal_eval(steps.strip())
                if isinstance(steps_list, list):
                    for s in steps_list:
                        if isinstance(s, dict) and 'action' in s and 'expected' in s:
                            steps_formatted.append(f"{s['action']} -> {s['expected']}")
                steps = '\n'.join(steps_formatted) if steps_formatted else steps.strip()
            except:
                steps = steps.strip()

            all_parsed_tcs.append({
                "User Story": tc_set["user_story"],
                "Test Type": test_type,
                "Title": title,
                "Priority": priority,
                "Steps": steps,
                "Status": "Not Executed",
                "Comments": ""
            })

    print(f"✓ Parsed {len(all_parsed_tcs)} test cases total")

    # Step 4: Compile all test cases for redundancy review
    print(f"\n🔍 Agent 2: Review Agent checking for redundant test cases...")

    compiled_tcs = ""
    for i, tc in enumerate(all_parsed_tcs, 1):
        compiled_tcs += f"\n{i}. Title: {tc['Title']}\n"
        compiled_tcs += f"   Steps:\n{tc['Steps']}\n"

    # Step 5: Call Review Agent to identify redundant test cases
    review_prompt = f"""You are a QA Review Lead Agent. Your ONLY job is to identify and remove duplicate/redundant/not needed/illogical test cases.

Here are all the generated test cases:

{compiled_tcs}

Instructions:
- Identify test cases that are duplicate redundant or not needed.
- **Identify only the main edge cases that add value and should be covered, removing any that are very similar or are too extreme and not practical.**
- **Only keep main test cases which cover the story and are very accurate and straightforward. Remove all useless test cases.**
- **Always remember QUALITY over QUANTITY.**
- Return a comma-separated list of test case numbers to KEEP 
- Only output the numbers, nothing else

Example output: 1,2,4,5,7,9,10,12

Output:"""

    try:
        review_response = openai.ChatCompletion.create(
            engine=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
            messages=[
                {"role": "system",
                 "content": "You are a QA Review Agent. Identify redundant test cases and return only the numbers to keep."},
                {"role": "user", "content": review_prompt}
            ],
            temperature=0.1,
            max_tokens=500
        )

        reviewed_content = review_response.choices[0].message.content.strip()

        # Parse the numbers to keep
        keep_indices = set()
        try:
            # Extract numbers from the response
            import re
            numbers = re.findall(r'\d+', reviewed_content)
            keep_indices = {int(n) for n in numbers if 1 <= int(n) <= len(all_parsed_tcs)}

            if not keep_indices:
                # If parsing failed, keep all
                keep_indices = set(range(1, len(all_parsed_tcs) + 1))
                print(f"⚠️ Could not parse review response, keeping all test cases")
            else:
                removed_count = len(all_parsed_tcs) - len(keep_indices)
                print(f"✅ Review completed - Removed {removed_count} redundant test cases")
        except Exception as parse_error:
            print(f"⚠️ Error parsing review response: {parse_error}, keeping all test cases")
            keep_indices = set(range(1, len(all_parsed_tcs) + 1))

        # Filter test cases based on review
        output = [tc for i, tc in enumerate(all_parsed_tcs, 1) if i in keep_indices]

    except Exception as e:
        print(f"❌ Review agent error: {e}")
        print(f"⚠️ Keeping all generated test cases")
        output = all_parsed_tcs

    # Step 6: Save final reviewed test cases
    df_out = pd.DataFrame(output)
    df_out.insert(0, "S.No.", range(1, len(df_out) + 1))
    df_out.to_excel(output_file, index=False)
    print(f"\n✅ Final output: {len(df_out)} reviewed test cases → {output_file}")
    return output_file


def upload_test_cases_ado(excel_file, org, proj, pat, plan_name, suite_name):
    df, mgr = pd.read_excel(excel_file), ADOTestManager(org, proj, pat, plan_name)
    upload_count, error_count = 0, 0
    print(f"🔄 Uploading to ADO suite '{suite_name}'...")
    for _, row in df.iterrows():
        if row.get("Status") == "Error": continue
        steps_str = str(row.get("Steps", "")).strip()
        try:
            # Parse steps from "action -> expected" format
            steps_list = []
            for line in steps_str.split('\n'):
                line = line.strip()
                if '->' in line:
                    parts = line.split('->', 1)
                    steps_list.append({'action': parts[0].strip(), 'expected': parts[1].strip()})

            if not steps_list:
                error_count += 1;
                continue

            mgr.create_test_case(suite_name=suite_name, title=row.get("Title", "Test Case"), steps=steps_list,
                                 priority=int(row.get("Priority", 2)))
            upload_count += 1
        except Exception as e:
            error_count += 1
            print(f"❌ {row.get('Title', 'Unknown')}: {str(e)}")
    print(f"✅ Uploaded {upload_count}/{len(df)} test cases ({error_count} failed)")
    return upload_count, error_count


def download_image(URL, FIGMA_TOKEN):
    pattern = r'figma\.com/(?:design|file)/([a-zA-Z0-9]+)/[^?]*\?node-id=([\d-]+)'
    match = re.search(pattern, URL)
    if not match:
        print("❌ Failed to parse Figma URL")
        return None
    FILE_KEY = match.group(1)
    node_id = match.group(2).replace('-', ':')
    headers = {
        "X-Figma-Token": FIGMA_TOKEN
    }
    img_api = f"https://api.figma.com/v1/images/{FILE_KEY}?ids={node_id}&format=png"
    img_response = requests.get(img_api, headers=headers).json()
    png_url = img_response["images"][node_id]
    image_bytes = requests.get(png_url).content
    with open("figma.png", "wb") as f:
        f.write(image_bytes)
    return "figma.png"
