---
name: info-gate
description: >-
  Information gate: ALWAYS invoke BEFORE generating any plan, program, scheme, or decision.
  Trigger keywords: 方案, 计划, plan, scheme, 制定, 设计, 规划, proposal, 路线图.
  Do NOT first judge whether information is "sufficient" — the gate's job IS to discover gaps.
  If any trigger keyword is present, stop and run the gate workflow (scan → ask → confirm → generate).
  NOT for: simple facts, one-off answers, spontaneous conversation.
  NOTE: Pair with a rule in RULE.md (see "info-gate" section) for guaranteed enforcement.
---

# Info Gate

Gatekeeper that ensures all necessary information is collected before generating output.
No plan leaves the gate until requirements are confirmed.

> **💡 建议配合 Rule 使用**：在 `RULE.md` 中添加 `info-gate` 章节，
> 确保 AI 在每次对话中都记得"遇方案先过 gate"。
> Skill 的触发依赖 agent 主动匹配关键词，Rule 提供强制兜底。

## Core Principle

> **Never guess. Gate first, generate second.**

When the user requests a plan, program, scheme, or decision, first scan for gaps.
If anything is missing, ambiguous, or conflicting — stop and ask. Do not proceed
with assumptions that could lead to mismatched output.

## Workflow

### Step 1 — Scan for Information Gaps

Before any plan generation, evaluate the request against these dimensions:

| Dimension | What to Check |
|-----------|--------------|
| **Goal** | What exactly should the output achieve? |
| **Constraints** | Time, budget, tools, environment, preferences |
| **Scope** | Boundaries — what's in, what's out |
| **Context** | Relevant history, existing conditions, prior decisions |
| **Format/Output** | How should the result be delivered? |
| **Priority** | What matters most when trade-offs exist? |

Identify gaps, ambiguities, and conflicts. Count them.

### Step 2 — Choose Mode by Gap Count

```
Gap count ≤ 5  →  Conversational mode: ask one question at a time
Gap count > 5  →  Questionnaire mode: output a Markdown form
```

If conversational, follow the one-at-a-time rule:
1. Ask the most important question first
2. Wait for response
3. Confirm understanding
4. Move to next question

### Step 3 — Questionnaire Format (for >5 gaps)

When outputting a questionnaire, use this exact format:

```markdown
## 📋 信息确认问卷

> 以下问题用于确保方案精确匹配你的需求。每个问题都有参考选项，
> 你可以：**选择选项** / **修改选项** / **自己填写**。

### Q1: [Question text]

- [ ] A) [Reference option — often the recommended one]
- [ ] B) [Alternative option]
- [ ] C) [Another alternative]
- [ ] ✏️ 自定义：_______

### Q2: [Question text]

- [ ] A) [Reference option]
- [ ] B) ...
- [ ] C) ...
- [ ] ✏️ 自定义：_______

---

> 回复示例：`Q1: A | Q2: 自定义 - xxxx | Q3: B | Q4: C`
```

**Rules for options:**
- Always provide 2-4 reference options per question
- Mark the recommended option with **(推荐)** when applicable
- Always include `✏️ 自定义` as the last option in every question
- Options must be mutually understandable — no jargon without explanation

### Step 4 — Process User Responses

When the user replies:

1. **Selected an option** → Use that as the answer
2. **Modified an option** → Use the modified version
3. **Wrote custom answer** → Use exactly what they wrote
4. **Selected multiple (valid for additive questions)** → Combine them

After processing, confirm the complete answer set before proceeding.

### Step 5 — Handle "Don't Know / Skip"

When the user says they don't know, "skip", "随便", or leaves blank:

1. **Provide a concrete recommendation** based on best judgment and context
2. **Note it clearly** with `💡 建议` marker so the user knows this was auto-filled
3. **Proceed** with that recommendation — do not loop indefinitely

```markdown
💡 建议：对于 Q3（训练频率），根据你的目标和时间安排，建议 **每周3天**。
   已按此选项继续。如需调整请告知。
```

**Never** leave a question hanging or block progress when user explicitly defers.

### Step 6 — Verify Completeness

Before generating any output, run a final check:

```
All gaps resolved?  →  Proceed to plan generation
Gaps still remain?  →  Return to Step 2 (re-count, re-mode)
```

Only when every necessary piece of information is confirmed (or deferred with a recommendation),
generate the plan.

### Step 7 — Generate

Generate the plan using ONLY the confirmed answers as source of truth.
Do not interpolate unselected options into the output.

## Iron Laws

1. **Gate before generate** — Never skip the scan. If you catch yourself guessing a requirement, stop and ask.
2. **≤5 conversational, >5 questionnaire** — Stick to the mode split. Don't dump 10 questions one-by-one.
3. **Every question has an escape hatch** — Always include `✏️ 自定义` and a skip path.
4. **Skip → recommend, note, proceed** — Do not stall. Provide a suggestion and move on.
5. **Output matches answers** — Generated content must derive from confirmed answers only, not from the full list of options presented.

## Anti-Patterns

| Anti-Pattern | Why It Fails | Correct Approach |
|---|---|---|
| Generating a plan with guessed assumptions | Output doesn't match user intent | Scan for gaps → ask → confirm → generate |
| Asking 8 questions one at a time | Exhausting, user disengages | Output a questionnaire when >5 |
| No custom option in questions | User's context may not fit any preset | Always include `✏️ 自定义` |
| Looping on "I don't know" | Stalls the session | Recommend and proceed with annotation |
| Using unselected option text in output | Plan includes things user didn't choose | Build from confirmed answers only |

## Examples

### Example 1: Conversational mode (≤5 gaps)

```
User: 帮我制定一个减脂计划

Agent: 好的！在制定之前先确认几个关键点。第一个问题：

       你的训练场地是？
       A) 健身房（有器械）
       B) 家里（只有哑铃）
       C) 家里（无器械纯自重）
       D) ✏️ 自定义

User: A

Agent: 明白，健身房。第二个问题：
       
       每周能投入几天训练？
       A) 2-3天
       B) 4-5天
       C) 6天以上
       D) ✏️ 自定义

User: B

Agent: 收到。信息完整，现在为你生成健身房4-5天/周的减脂计划...
```

### Example 2: Questionnaire mode (>5 gaps)

```
User: 帮我设计一个完整的增肌方案

Agent: （扫描发现8个信息缺口）→ 输出 Markdown 问卷

## 📋 信息确认问卷

Q1: 训练场地？
- [ ] A) 健身房
- [ ] B) 家里（哑铃为主）
- [ ] C) 家里（纯自重）
- [ ] ✏️ 自定义

Q2: 每周训练天数？
- [ ] A) 2-3天
- [ ] B) 4-5天
- [ ] C) 6天以上
- [ ] ✏️ 自定义

...（其余问题）

User: Q1: A | Q2: B | Q3: 不知道 | Q4: 自定义 - 希望加入瑜伽 | ...

Agent: 收到。Q3你选择了「不知道」，💡 建议：根据已有信息，推荐 **PPL分化（推/拉/腿）**。
       确认无误的话我按全部确认项生成方案。
```

### Example 3: Skip handling

```
User: Q5: 不知道，随便吧

Agent: 💡 建议：对于 Q5（蛋白质摄入目标），根据你的体重和目标，建议 **每天150g**。
       已按此选项继续，如需调整请告知。
```
