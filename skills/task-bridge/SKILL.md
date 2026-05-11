---
name: task-bridge
description: >-
  Cross-session task bridge — saves a self-contained instruction to memory that can be
  executed in a future session without re-describing the task. The instruction is
  automatically deleted after execution to avoid memory pollution.
  Use when: (1) User says "跨对话任务", "下次对话执行", "跨session任务",
  (2) User wants to defer a task to a future conversation,
  (3) User says "执行待完成任务" or "执行待办指令" in any session.
---

# Task Bridge

Bridge tasks across sessions. Save a self-contained instruction now,
execute it in any future session, then auto-clean.

## Two-Phase Workflow

### Phase 1 — Save (current session)

When the user says they want a cross-session task:

1. **Understand the task** — Ask clarifying questions until the instruction is complete and self-contained.
   A future agent instance (with no context of this conversation) must be able to execute it solely
   from the stored text.

2. **Format the instruction** — Use this exact format in MEMORY.md:

```markdown
## 🔄 待执行指令：<简短标题>
> 仅一次，执行后立即擦除 | 创建于 YYYY-MM-DD

<完整的自包含指令，含所有必要上下文>

---
```

3. **Quality check** — Before saving, verify: if you were a fresh agent reading this cold,
   would you understand exactly what to do, how to do it, and what "done" looks like?

4. **Save** — Append to MEMORY.md using the `edit` tool.

### Phase 2 — Execute (future session)

When the user says "执行待完成任务" / "执行待办指令" / "run pending task" etc.:

1. **Read MEMORY.md** — Scan for `## 🔄 待执行指令：` marker
2. **If found**:
   - Read the full instruction
   - Execute it immediately
   - After completion, **delete the entire block** from MEMORY.md (from `## 🔄` through `---`)
   - Report: "✅ 指令「<标题>」已执行完毕并擦除。结果：..."
3. **If not found** — Reply: "没有找到待执行指令。"
4. **If multiple found** — List them and ask which to execute.

## Instruction Writing Guidelines

A good cross-session instruction must be:

| Trait | Check |
|-------|-------|
| **Self-contained** | No references to "what we discussed earlier" — everything spelled out |
| **Actionable** | Clear what to do, not just what to think about |
| **Scoped** | Has a clear "done" condition |
| **Concrete** | Specific files, commands, parameters — not vague descriptions |

### Good vs Bad

```
❌ BAD: 帮我把上次那个数据分析做了
   → "上次" is undefined in a new session

✅ GOOD: 分析 C:\data\experiment\results.csv：
         (1) 计算每列均值和标准差
         (2) 找出异常值（>3σ）
         (3) 输出 Markdown 报告到 C:\data\experiment\report.md
```

## Storage Location

Instructions are stored in **MEMORY.md** (the long-term memory index), because:
- MEMORY.md is auto-loaded at session start → immediately visible
- Single known location → no searching needed
- Easy cleanup with edit tool

## Cleanup Rule (Iron Law)

> **Delete immediately after execution. No exceptions.**

Once a task is executed, remove its entire block from MEMORY.md.
This prevents:
- Re-executing completed tasks
- Memory index bloat
- Confusion between active and completed tasks

## Interaction with Other Skills

If the pending instruction involves a skill (e.g., workout-program-designer),
the future agent should invoke that skill normally during execution.
The instruction should mention the skill name if relevant.

## Examples

### Saving a task

```
User: 我要做一个跨对话任务，下次帮我在knowledge里建一个航空发动机叶片的页面

Agent: 好的，让我确认一下这个跨对话任务的内容。

       你需要我下次对话时在 knowledge/ 中创建一个关于航空发动机叶片的页面。
       请问：
       - 资料来源是什么？（你现在有文章/链接吗？还是通用知识？）
       - 你希望侧重哪个方面？（设计原理 / 失效模式 / 材料 / 制造工艺？）

User: 通用知识，侧重失效模式和设计原理

Agent: 明白。已存入跨对话指令。

## 🔄 待执行指令：创建航空发动机叶片知识页
> 仅一次，执行后立即擦除 | 创建于 2026-05-11

在 knowledge/concepts/ 下创建 turbine-blade.md，内容覆盖：
1. 涡轮叶片失效模式（蠕变、热疲劳、氧化、外物损伤）
2. 设计原理（冷却结构、气动外形、榫头连接）
来源为通用工程知识，不需要外部文章。
创建后同步更新 knowledge/index.md。
---
```

### Executing a task

```
User: 执行待完成任务

Agent: （读取 MEMORY.md，找到指令，执行...）

       已创建 knowledge/concepts/turbine-blade.md ✅
       已更新 knowledge/index.md ✅

       ✅ 指令「创建航空发动机叶片知识页」已执行完毕并擦除。
```
