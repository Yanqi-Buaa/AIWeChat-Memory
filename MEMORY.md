## 个人设定
- 用户称呼YQ，助手命名小D；交流风格实事求是，科研知识专业严谨，日常健身放松并给予情绪价值
## 系统架构
- 记忆系统跨对话持久保存，分为三层：即时、每日（memory/ 目录）和知识库（memory/long-term/ 目录），每次对话自动加载 MEMORY.md
## 技能配置
- 已创建 info-gate 技能（2026-05-11）：通用信息闸门，制定方案前扫描信息缺口，≤5题逐题追问，>5题输出Markdown问卷含参考选项，用户可选/改/自定义，跳过时给建议选项并注明执行
- 已创建 task-bridge 技能（2026-05-11）：跨对话任务桥接，保存自包含指令到 MEMORY.md（格式 `## 🔄 待执行指令：<标题>`），下次对话说「执行待完成任务」即可执行，完成后自动擦除
- 已安装5个skill：knowledge-wiki、skill-creator、find-skills、workout-program-designer、image-generation（缺API key未配置）
- workout-program-designer技能可定制训练计划，涵盖目标定制、渐进超负荷、休息日优化、场景适配（家庭/健身房）、减载周、损伤预防和进度追踪
## 健身
- 健身档案与训练日志已迁移至 [knowledge/fitness/](knowledge/fitness/) → [档案](knowledge/fitness/profile.md) | [日志](knowledge/fitness/log.md)
- **当前总次数**: 第31次（2026-05-11 拉日）
- **第32次**: 5月12日周二（腿日）
- **5月14日周四起**: 正式启动四模板轮转计划（拉→推A→腿→推B），第33次开始
- **记录要求**: 每次训练后汇报数据，需标注第几次训练
## 定时简报系统（2026-05-11 设计）
- **架构**: config/topics.md 为话题注册表，每个话题含轻量简报+深度简报两套自包含prompt，可随时增删改话题
- **话题-1 [AI前沿]**: 轻量简报 Tue-Sun 早7:00，深度简报 Mon 早7:00
- **调度方式**: 用户自行在scheduler中创建任务，从config/topics.md复制对应prompt到ai_task参数
- **扩展**: 新增话题只需在config/topics.md照模板添加，再创建对应scheduler任务即可


---
