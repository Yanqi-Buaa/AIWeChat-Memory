## 个人设定
- 用户称呼YQ，助手命名小D；交流风格实事求是，科研知识专业严谨，日常健身放松并给予情绪价值
## 系统架构
- 记忆系统跨对话持久保存，分为三层：即时、每日（memory/ 目录）和知识库（memory/long-term/ 目录），每次对话自动加载 MEMORY.md
## 健身
- 健身档案与训练日志已迁移至 knowledge/fitness/（档案profile.md，日志log.md）
- 当前总次数：第35次（2026-05-17 腿日 ✓，107.5kg×5×3，1RM路线#1），下次第36次推B日（肩）
- 每次训练后汇报数据，需标注第几次训练
## 职业规划
- 博士方向：转子动力学/结构强度 → 数据挖掘/故障智能诊断（组内第一个开此方向）
- 技能定位：组内软件编程+AI工具最擅长且最了解需求的人
- 毕业后意向：不打算留校，大概率进航天企业（CASC/CASIC），同时以"编外人员"身份与课题组长期合作
- 5年策略：先小范围与关系最好的师兄合作，等大老板和二老板退休（约5年）后扩大范围
- 收入通道：有朋友已开公司，可作为外部合同方与课题组签约；收入归属需规避航天企业在职兼职限制
- 详细分析见 [knowledge/analysis/career-external-collaboration.md](knowledge/analysis/career-external-collaboration.md)
## 定时简报系统（2026-05-11 设计）
- 架构：config/topics.md 为话题注册表，每个话题含轻量简报+深度简报两套自包含prompt，可随时增删改话题
- 话题-1 [AI前沿]：轻量简报 Tue-Sun 早7:00，深度简报 Mon 早7:00
- 调度方式：用户自行在scheduler中创建任务，从config/topics.md复制对应prompt到ai_task参数
- 扩展：新增话题只需在config/topics.md照模板添加，再创建对应scheduler任务即可
## 重启流程
当YQ说「重启」「cow restart」「cow start」时：
1. 先确认旧 cmd 窗口已关闭 — 否则 `cow start` 报 `Permission denied: nohup.out`（旧进程占住文件）
2. 在命令行执行 `cow start`
3. 如需发重启通知 → app.py 启动时读 `pending_notify.json` → 写 `scripts/.delayed_notify.txt` → YQ发一条消息后自动送达
4. 看门狗托管重启：写 `pending_notify.json` → 创建一次性 schtasks（30s 后触发）→ 自杀 → 看门狗兜底
5. ⚠️ 看门狗已修复自毁 bug：watchdog.bat 末尾执行 `schtasks /Delete`，跑完不再弹窗
## 项目特性
- 已实现功能、待开发功能、技能安装记录等均记录在 `FEATURES.md`
- 规则：每次实现新功能或讨论功能构想时，同步更新 FEATURES.md
## 服务器信息
- 主节点：`ssh yanqi@172.17.135.116` → 再从主节点 `ssh -X gpu1`（gpu1 无独立IP）
- 工作目录：`/data/yanqi/PINNs/`
- tmux 会话名：`deepseek`（在 gpu1 上）
- 任务简报：用 `find` 定位 `runtime_brief*` 文件，不固定在单一路径
