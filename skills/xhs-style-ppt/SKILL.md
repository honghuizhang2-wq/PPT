---
name: xhs-style-ppt
description: "Create Xiaohongshu (小红书) style PPT presentations by analyzing reference images from XHS creators and generating slides that mimic their visual style. Use when: (1) user provides XHS screenshots/images and asks to make a PPT in that style; (2) user mentions 小红书 style PPT/幻灯片; (3) user wants to create content cards, guide posts, or tip-style presentations with XHS aesthetics (warm colors, hand-drawn feel, sticky notes, emoji decorations, card layouts)."
---

# XHS Style PPT Creator

Create PPT presentations that mimic the visual style of Xiaohongshu (小红书) content creators.

## Workflow

### Step 1: Analyze Reference Images

When user provides XHS screenshots or creator links:

1. **Extract creator profile** from the link (name, bio, tags, niche)
2. **Analyze images** using mimo-omni skill (`bash mimo_api.sh image <path> "分析视觉风格"`) to identify:
   - Color palette (main/accent/background colors)
   - Typography (title style, body font feel)
   - Layout pattern (cards, columns, sections, sticky notes)
   - Illustration style (hand-drawn, flat, photo-based)
   - Content structure (总-分-总, list, Q&A, tip-based)
   - Emotional tone (治愈, 活泼, 高级感, etc.)

If no images provided, use the default XHS style profile from [references/style_patterns.md](references/style_patterns.md).

### Step 2: Plan Content Structure

Based on analysis, plan 6-10 slides with:

- **Cover slide**: Title + subtitle + decorative elements
- **Content slides**: Card-based layout with emoji icons, 2-column grid
- **Quote/tip slides**: Sticky-note style or highlight blocks
- **Closing slide**: Emotional message + call-to-action

Content rules:
- Use conversational, warm tone (not lecture-style)
- Include emoji as visual anchors for each point
- Keep text concise: one idea per card
- Use relatable examples and self-deprecating humor
- End with emotional resonance or actionable takeaway

### Step 3: Generate PPT

Run the bundled script to create the PPT:

```bash
python3 ~/.openclaw/skills/xhs-style-ppt/scripts/create_xhs_ppt.py \
  --title "主题标题" \
  --subtitle "副标题" \
  --output output.pptx \
  --palette warm \
  --slides-json slides.json
```

**Palette options**: `warm` (default), `cool`, `pastel`, `earth`

**slides.json format**:
```json
[
  {
    "type": "cover",
    "title": "主标题",
    "subtitle": "副标题"
  },
  {
    "type": "cards",
    "section_title": "📋 章节标题",
    "cards": [
      {"emoji": "😰", "title": "卡片标题", "desc": "卡片描述\n支持换行"},
      {"emoji": "💡", "title": "卡片标题", "desc": "卡片描述"}
    ]
  },
  {
    "type": "tips",
    "section_title": "🌿 章节标题",
    "tips": [
      {"num": "1️⃣", "title": "建议标题", "desc": "建议内容"},
      {"num": "2️⃣", "title": "建议标题", "desc": "建议内容"}
    ]
  },
  {
    "type": "quotes",
    "section_title": "💬 语录标题",
    "quotes": ["「第一句金句」", "「第二句金句」"]
  },
  {
    "type": "message",
    "title": "🌱 结尾标题",
    "body": "长文本内容，支持\\n换行"
  },
  {
    "type": "closing",
    "title": "主题文字",
    "tagline": "结语",
    "cta": "收藏引导文案"
  }
]
```

### Step 4: QA & Iterate

After generation:
1. Check slide count and content completeness
2. Verify color consistency across slides
3. Adjust text length if cards overflow
4. User can request style changes (palette, layout density, tone)

## Style Reference

See [references/style_patterns.md](references/style_patterns.md) for detailed XHS visual style patterns and color palettes.

## Dependencies

- `python-pptx` (install: `pip3 install --break-system-packages python-pptx`)
- `mimo-omni` skill (for image analysis, optional)
