## 个人设定
- 用户称呼YQ，助手命名小D；交流风格实事求是，科研知识专业严谨，日常健身放松并给予情绪价值
## 系统架构
- 记忆系统跨对话持久保存，分为三层：即时、每日（memory/ 目录）和知识库（memory/long-term/ 目录），每次对话自动加载 MEMORY.md

## 健身
- 健身档案与训练日志已迁移至 [knowledge/fitness/](knowledge/fitness/) → [档案](knowledge/fitness/profile.md) | [日志](knowledge/fitness/log.md)
- **当前总次数**: 第32次（2026-05-12 腿日）
- **第33次**: 5月14日周四（拉日——四模板轮转首日）
- **5月14日周四起**: 正式启动四模板轮转计划（拉→推A→腿→推B），第33次开始
- **记录要求**: 每次训练后汇报数据，需标注第几次训练
## 项目管理
- 师弟协作项目档案存放于 `projects/` 目录，与个人记忆/知识库隔离，具体内容直接读该目录文件
## 定时简报系统（2026-05-11 设计）
- **架构**: config/topics.md 为话题注册表，每个话题含轻量简报+深度简报两套自包含prompt，可随时增删改话题
- **话题-1 [AI前沿]**: 轻量简报 Tue-Sun 早7:00，深度简报 Mon 早7:00
- **调度方式**: 用户自行在scheduler中创建任务，从config/topics.md复制对应prompt到ai_task参数
- **扩展**: 新增话题只需在config/topics.md照模板添加，再创建对应scheduler任务即可
## 项目特性
- 已实现功能、待开发功能、技能安装记录等均记录在 `A:\AIWeChat\CowAgent\FEATURES.md`
- **规则**: 每次实现新功能或讨论功能构想时，同步更新 FEATURES.md
