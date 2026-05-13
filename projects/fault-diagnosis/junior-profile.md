# 师弟档案

## 基本信息

| 维度 | 初始状态 | 当前状态 |
|------|----------|----------|
| Python | 刚入门，可借助AI辅助编程 | 已能借助 Cursor/AI 完成完整项目 |
| 数据处理 | 用过 NumPy/Pandas，不熟练 | 有实战经验（预处理 .mat 数据、train/val/test 划分） |
| 数学 | 本科水平，大部分已遗忘 | — |
| 深度学习理论 | 理解前向/反向传播，未动手实践 | 已动手实践 CNN 分类 |
| PyTorch | 未上手 | 已能搭建训练流程（DataLoader、训练循环、checkpoint） |
| 信号处理 | 未涉及 | 未涉及 |
| 可用时间 | 每周 5~10 小时 | — |

## 工具能力

- ✅ 使用 Cursor + AI 辅助编程
- ✅ PyTorch 项目结构搭建（config YAML、实验管理、模型对比）
- ✅ 虚拟环境管理（venv）

## 当前任务

- **进行中**：多标签频谱分类（Multi-label ComparisonNet）
  - 技术点：BCEWithLogitsLoss / Hamming Loss / 复用 Comparator
  - 启动时间：2026-05-13
  - 状态：等待交付

## 已完成项目

1. **Iris 分类练手** — 入门级，熟悉基本流程
2. **频谱分类实验** — 2D CNN 频谱分类，搭建完整训练管线，含双模型对比
   - ComparisonNet2D：test_acc 0.9955，参数量 381
   - SimpleCNN2D Baseline：test_acc 0.9911，参数量 5508
   - 结论：轻量模型精度反超，项目结构规范
