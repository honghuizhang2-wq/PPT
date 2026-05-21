# XHS Style PPT

小红书风格 PPT 生成工具 + 示例

## 📁 目录结构

```
├── output/
│   └── 东亚人类反焦虑平静指南.pptx   # 示例 PPT（8页，9:16竖版）
├── skills/
│   ├── xhs-style-ppt/                 # OpenClaw Skill 源码
│   │   ├── SKILL.md                   # 技能说明与工作流
│   │   ├── scripts/
│   │   │   └── create_xhs_ppt.py      # PPT 生成脚本
│   │   └── references/
│   │       ├── style_patterns.md      # XHS 视觉风格参考
│   │       └── slides_example.json    # 示例配置
│   └── xhs-style-ppt.skill           # 打包好的技能文件
└── README.md
```

## 🚀 使用方式

### 作为 OpenClaw Skill 使用

将 `skills/xhs-style-ppt/` 目录复制到 `~/.openclaw/skills/` 即可。

### 命令行直接使用

```bash
python3 scripts/create_xhs_ppt.py \
  --slides-json slides.json \
  --output my_ppt.pptx \
  --palette warm
```

**配色方案**: `warm`(治愈系) / `cool`(清新系) / `pastel`(马卡龙) / `earth`(大地系)

### slides.json 格式

支持 6 种幻灯片类型：

| 类型 | 说明 |
|------|------|
| `cover` | 封面页 |
| `cards` | 卡片网格（2列） |
| `tips` | 编号建议列表 |
| `quotes` | 便签风格金句 |
| `message` | 长文情感页 |
| `closing` | 结尾页 |

## 🎨 设计风格

模仿小红书内容创作者的视觉语言：

- 🎨 暖色治愈系配色（绿/黄/粉/紫/橙交替）
- 📐 圆角卡片 + 彩色色块分区排版
- 📝 黄色便签纸风格的金句区
- 🌱 emoji 装饰替代手绘插画
- 💬 口语化、温暖、不教条的文案语气

## 依赖

```bash
pip install python-pptx
```
