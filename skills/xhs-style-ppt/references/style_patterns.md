# XHS Visual Style Patterns

Reference for mimicking Xiaohongshu content creator visual styles.

## Common XHS Style Archetypes

### 1. 治愈手绘风 (Healing Hand-drawn)
- **Palette**: Soft green, warm yellow, cream white, blush pink
- **Layout**: Vertical sections with color blocks, sticky-note accents
- **Typography**: Bold cartoon titles, clean body text
- **Illustrations**: Cute abstract characters, simple line art
- **Tone**: Warm, empathetic, self-deprecating humor
- **Best for**: Anxiety guides, self-care tips, emotional content

### 2. 极简高级感 (Minimalist Premium)
- **Palette**: Black, white, beige, gold accents
- **Layout**: Lots of whitespace, single-column, large typography
- **Typography**: Serif or elegant sans-serif, generous line spacing
- **Illustrations**: None or minimal geometric shapes
- **Tone**: Sophisticated, calm, authoritative
- **Best for**: Fashion, lifestyle, knowledge sharing

### 3. 活泼可爱风 (Playful Cute)
- **Palette**: Bright pink, baby blue, mint, lavender
- **Layout**: Grid cards, rounded corners, playful alignment
- **Typography**: Rounded fonts, mixed sizes, emoji-heavy
- **Illustrations**: Sticker-style, kawaii characters
- **Tone**: Energetic, fun, Gen-Z vibe
- **Best for**: Food, travel, daily life, recommendations

### 4. 干货知识风 (Knowledge/Value)
- **Palette**: Navy, white, accent yellow or coral
- **Layout**: Numbered lists, two-column comparison, step-by-step
- **Typography**: Bold section headers, clean numbered items
- **Illustrations**: Icons, charts, simple diagrams
- **Tone**: Direct, practical, "save this post" energy
- **Best for**: Tutorials, how-to guides, professional tips

## Color Palettes (Hex Values)

### Warm (治愈系)
```python
BG_CREAM    = "#FDF5E6"   # 奶油白底
BG_GREEN    = "#D4EDDA"   # 薄荷绿
BG_YELLOW   = "#FFF3CD"   # 暖黄色
BG_PINK     = "#FCE4EC"   # 淡粉色
BG_LAVENDER = "#E8DAEF"   # 淡紫色
BG_PEACH    = "#FFE5D0"   # 淡橙色
BG_MINT     = "#D1F2EB"   # 薄荷色
STICKY_YEL  = "#FFF9C4"   # 便签黄
TEXT_DARK   = "#2D2D2D"   # 深色文字
TEXT_BROWN  = "#5D4E37"   # 棕色文字
TEXT_GREEN  = "#2E7D32"   # 绿色文字
```

### Cool (清新系)
```python
BG_ICE      = "#E8F4FD"   # 冰蓝底
BG_SKY      = "#B3E5FC"   # 天蓝
BG_MINT     = "#C8E6C9"   # 薄荷
BG_LILAC    = "#F3E5F5"   # 丁香
TEXT_NAVY    = "#1A237E"   # 深蓝文字
TEXT_SLATE   = "#37474F"   # 石板灰
```

### Pastel (马卡龙系)
```python
BG_ROSE     = "#F8BBD0"   # 玫瑰粉
BG_PEACH    = "#FFCCBC"   # 蜜桃
BG_LEMON    = "#FFF9C4"   # 柠檬
BG_LAVENDER = "#D1C4E9"   # 薰衣草
BG_MINT     = "#C8E6C9"   # 薄荷
TEXT_PLUM    = "#4A148C"   # 梅子色文字
```

### Earth (大地系)
```python
BG_SAND     = "#F5E6CC"   # 沙色
BG_SAGE     = "#D5DBDB"   # 鼠尾草
BG_TERRA    = "#E8D5B7"   # 赤陶
BG_OLIVE    = "#D5E8D4"   # 橄榄
TEXT_BROWN   = "#3E2723"   # 深棕文字
TEXT_OLIVE   = "#33691E"   # 橄榄绿文字
```

## Layout Patterns

### Two-Column Card Grid
```
┌─────────┐  ┌─────────┐
│  📌     │  │  💡     │
│ Title   │  │ Title   │
│ Desc    │  │ Desc    │
└─────────┘  └─────────┘
┌─────────┐  ┌─────────┐
│  🎯     │  │  🌟     │
│ Title   │  │ Title   │
│ Desc    │  │ Desc    │
└─────────┘  └─────────┘
```

### Sticky Note Quotes
```
┌───────────────────┐
│  💛 「金句内容」    │  ← Yellow sticky background
│                   │
└───────────────────┘
```

### Section Title + Cards
```
┌──────────────────────────────┐
│     📋 章节大标题              │  ← Section header
├──────────────────────────────┤
│ ┌────────┐ ┌────────┐       │
│ │ Card 1 │ │ Card 2 │       │
│ └────────┘ └────────┘       │
│ ┌────────┐ ┌────────┐       │
│ │ Card 3 │ │ Card 4 │       │
│ └────────┘ └────────┘       │
└──────────────────────────────┘
```

### Left-Icon Right-Content
```
┌───┐  ┌──────────────────┐
│   │  │ 大标题            │
│ 🏫│  │ 详细描述文字       │
│   │  │ 可以多行           │
└───┘  └──────────────────┘
```

## Typography Guidelines

- **Titles**: 32-48pt, bold, centered or left-aligned
- **Section headers**: 22-36pt, bold
- **Card titles**: 16-18pt, bold
- **Body text**: 12-14pt, regular, line spacing 1.5-1.8x
- **Sticky notes**: 14-17pt, centered
- **Font**: Microsoft YaHei (微软雅黑) for Chinese, fallback to system sans-serif

## Content Writing Style

### Tone Rules
1. **对话感** > 说教感: Write like talking to a friend
2. **自嘲** > 炫耀: Self-deprecating humor builds trust
3. **共鸣** > 教导: "你是不是也这样" > "你应该这样做"
4. **具体** > 抽象: "奶茶涨价速度" > "物价上涨"
5. **短句** > 长句: One thought per line, max 2 lines per card

### Title Formulas
- 「东亚人类XXX指南」
- 「给XX一些XXX的小tips」
- 「如果你也XXX，这篇是写给你的」
- 「XXX的N个真相」
- 「允许自己XXX」

### Closing Patterns
- Emotional letter format (亲爱的...)
- Permission-giving statement (你可以...)
- CTA: 「觉得有用就收藏吧💛」
