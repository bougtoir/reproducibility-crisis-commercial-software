#!/usr/bin/env python3
"""Create PPTX files with one figure per slide (English and Japanese versions)."""

import json
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

OUTPUT_DIR = Path("/home/ubuntu/reproducibility_study/output")
FIG_DIR = OUTPUT_DIR / "figures"

with open(OUTPUT_DIR / "summary_stats.json") as f:
    stats = json.load(f)
overall = stats["overall"]
N = overall["total_papers"]

# Slide dimensions: widescreen 13.333 x 7.5 inches
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ── Figure definitions ──────────────────────────────────────────────────

figures_en = [
    {
        "file": "fig1_software_rates_by_field.png",
        "title": "Figure 1. Software Mention Rates by Research Field",
        "caption": f"Grouped bar chart showing the percentage of papers mentioning any software, commercial software only, and open-source software only, across seven research fields. N={N:,} papers.",
    },
    {
        "file": "fig2_top_commercial_software.png",
        "title": "Figure 2. Most Frequently Used Commercial Software",
        "caption": f"Horizontal bar chart of the top 15 commercial software tools by frequency of citation. N={N:,} papers.",
    },
    {
        "file": "fig3_version_and_availability.png",
        "title": "Figure 3. Version Reporting and Data/Code Availability",
        "caption": "(a) Version mention rate among papers citing software, by research field. (b) Code and data availability statement rates by research field.",
    },
    {
        "file": "fig4_replication_costs.png",
        "title": "Figure 4. Estimated Software Replication Costs",
        "caption": "(a) Distribution of estimated replication costs among papers using commercial software. (b) Mean replication cost by research field.",
    },
    {
        "file": "fig5_software_landscape.png",
        "title": "Figure 5. Top 20 Software Tools in Published Research",
        "caption": f"Horizontal bar chart of top 20 software tools colored by license type (red = commercial, green = open-source). N={N:,} papers.",
    },
    {
        "file": "fig6_version_availability.png",
        "title": "Figure 6. Version Availability of Commercial Software",
        "caption": "Pie chart showing the availability status of commercial software versions cited in published papers. Legend indicates counts and percentages.",
    },
    {
        "file": "fig7_software_heatmap.png",
        "title": "Figure 7. Software Usage Across Research Fields",
        "caption": "Heatmap of top 15 software tools across seven research fields, showing field-specific software usage patterns.",
    },
    {
        "file": "fig8_pmc_impact.png",
        "title": "Figure 8. Impact of PMC Full-Text on Software Detection",
        "caption": "(a) PMC full-text availability rate by research field. (b) Comparison of software detection rates with and without PMC full-text access.",
    },
]

figures_jp = [
    {
        "file": "fig1_software_rates_by_field.png",
        "title": "\u56f31. \u7814\u7a76\u5206\u91ce\u5225\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u8a00\u53ca\u7387",
        "caption": f"\u5168\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u3001\u5546\u7528\u306e\u307f\u3001\u30aa\u30fc\u30d7\u30f3\u30bd\u30fc\u30b9\u306e\u307f\u306e3\u30ab\u30c6\u30b4\u30ea\u3067\u30017\u7814\u7a76\u5206\u91ce\u306b\u304a\u3051\u308b\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u8a00\u53ca\u7387\u3092\u793a\u3059\u30b0\u30eb\u30fc\u30d7\u30d0\u30fc\u30c1\u30e3\u30fc\u30c8\u3002N={N:,}\u672c\u3002",
    },
    {
        "file": "fig2_top_commercial_software.png",
        "title": "\u56f32. \u6700\u9812\u51fa\u5546\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2",
        "caption": f"\u5f15\u7528\u983b\u5ea6\u4e0a\u4f4d15\u306e\u5546\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u30c4\u30fc\u30eb\u306e\u6a2a\u68d2\u30b0\u30e9\u30d5\u3002N={N:,}\u672c\u3002",
    },
    {
        "file": "fig3_version_and_availability.png",
        "title": "\u56f33. \u30d0\u30fc\u30b8\u30e7\u30f3\u5831\u544a\u3068\u30c7\u30fc\u30bf/\u30b3\u30fc\u30c9\u5229\u7528\u53ef\u80fd\u6027",
        "caption": "(a) \u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u3092\u5f15\u7528\u3057\u305f\u8ad6\u6587\u306b\u304a\u3051\u308b\u30d0\u30fc\u30b8\u30e7\u30f3\u8a00\u53ca\u7387\uff08\u5206\u91ce\u5225\uff09\u3002(b) \u30b3\u30fc\u30c9\u304a\u3088\u3073\u30c7\u30fc\u30bf\u5229\u7528\u53ef\u80fd\u6027\u58f0\u660e\u7387\uff08\u5206\u91ce\u5225\uff09\u3002",
    },
    {
        "file": "fig4_replication_costs.png",
        "title": "\u56f34. \u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u8ffd\u8a66\u30b3\u30b9\u30c8\u306e\u63a8\u5b9a",
        "caption": "(a) \u5546\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u4f7f\u7528\u8ad6\u6587\u306e\u8ffd\u8a66\u30b3\u30b9\u30c8\u5206\u5e03\u3002(b) \u7814\u7a76\u5206\u91ce\u5225\u5e73\u5747\u8ffd\u8a66\u30b3\u30b9\u30c8\u3002",
    },
    {
        "file": "fig5_software_landscape.png",
        "title": "\u56f35. \u5b66\u8853\u8ad6\u6587\u306b\u304a\u3051\u308b\u4e0a\u4f4d20\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2",
        "caption": f"\u30e9\u30a4\u30bb\u30f3\u30b9\u7a2e\u5225\u3067\u8272\u5206\u3051\u3057\u305f\u4e0a\u4f4d20\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\uff08\u8d64=\u5546\u7528\u3001\u7dd1=\u30aa\u30fc\u30d7\u30f3\u30bd\u30fc\u30b9\uff09\u3002N={N:,}\u672c\u3002",
    },
    {
        "file": "fig6_version_availability.png",
        "title": "\u56f36. \u5546\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u30d0\u30fc\u30b8\u30e7\u30f3\u306e\u5165\u624b\u53ef\u80fd\u6027",
        "caption": "\u8ad6\u6587\u3067\u5f15\u7528\u3055\u308c\u305f\u5546\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u30d0\u30fc\u30b8\u30e7\u30f3\u306e\u5165\u624b\u53ef\u80fd\u6027\u3092\u793a\u3059\u30d1\u30a4\u30c1\u30e3\u30fc\u30c8\u3002\u51e1\u4f8b\u306b\u4ef6\u6570\u3068\u5272\u5408\u3092\u8868\u793a\u3002",
    },
    {
        "file": "fig7_software_heatmap.png",
        "title": "\u56f37. \u7814\u7a76\u5206\u91ce\u6a2a\u65ad\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u4f7f\u7528\u30d2\u30fc\u30c8\u30de\u30c3\u30d7",
        "caption": "7\u7814\u7a76\u5206\u91ce\u306b\u304a\u3051\u308b\u4e0a\u4f4d15\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u306e\u4f7f\u7528\u5206\u5e03\u30d2\u30fc\u30c8\u30de\u30c3\u30d7\u3002\u5206\u91ce\u56fa\u6709\u306e\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u4f7f\u7528\u30d1\u30bf\u30fc\u30f3\u3092\u793a\u3059\u3002",
    },
    {
        "file": "fig8_pmc_impact.png",
        "title": "\u56f38. PMC\u5168\u6587\u30a2\u30af\u30bb\u30b9\u304c\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u691c\u51fa\u306b\u4e0e\u3048\u308b\u5f71\u97ff",
        "caption": "(a) \u7814\u7a76\u5206\u91ce\u5225PMC\u5168\u6587\u5229\u7528\u53ef\u80fd\u7387\u3002(b) PMC\u5168\u6587\u306e\u6709\u7121\u306b\u3088\u308b\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u691c\u51fa\u7387\u306e\u6bd4\u8f03\u3002",
    },
]


def create_pptx(figures, lang, title_text):
    """Create a PPTX with one figure per slide."""
    prs = Presentation()
    # Set widescreen 16:9
    prs.slide_width = Emu(12192000)   # 13.333 inches
    prs.slide_height = Emu(6858000)   # 7.5 inches

    slide_w = 12192000  # EMU
    slide_h = 6858000

    # ── Title slide ──
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Title text box
    from pptx.util import Emu as E
    txBox = slide.shapes.add_textbox(
        Inches(1), Inches(2), Inches(11.333), Inches(2)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x1A, 0x23, 0x7E)
    p.alignment = PP_ALIGN.CENTER

    # Subtitle
    p2 = tf.add_paragraph()
    if lang == "en":
        p2.text = f"N={N:,} papers | PubMed Stratified Random Sampling (2020\u20132025)"
    else:
        p2.text = f"N={N:,}\u672c | PubMed\u5c64\u5225\u30e9\u30f3\u30c0\u30e0\u30b5\u30f3\u30d7\u30ea\u30f3\u30b0\uff082020\u20132025\uff09"
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(0x42, 0x42, 0x42)
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(20)

    # ── Figure slides ──
    for fig_info in figures:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank

        fig_path = FIG_DIR / fig_info["file"]
        title = fig_info["title"]
        caption = fig_info["caption"]

        # Title at top
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.2), Inches(12.333), Inches(0.7)
        )
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(0x1A, 0x23, 0x7E)
        p_title.alignment = PP_ALIGN.CENTER

        # Figure image - centered, large
        # Available area: ~12 x 5.5 inches for the image
        max_img_w = Inches(11)
        max_img_h = Inches(5.3)

        # Get image dimensions to maintain aspect ratio
        from PIL import Image
        with Image.open(fig_path) as img:
            img_w, img_h = img.size
        
        aspect = img_w / img_h
        # Fit within max_img_w x max_img_h
        if aspect > (11 / 5.3):
            # Width-limited
            pic_w = max_img_w
            pic_h = int(max_img_w / aspect)
        else:
            # Height-limited
            pic_h = max_img_h
            pic_w = int(max_img_h * aspect)

        # Center horizontally
        left = int((slide_w - pic_w) / 2)
        top = Inches(1.0)

        slide.shapes.add_picture(str(fig_path), left, top, pic_w, pic_h)

        # Caption at bottom
        caption_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(6.5), Inches(11.733), Inches(0.8)
        )
        tf_cap = caption_box.text_frame
        tf_cap.word_wrap = True
        p_cap = tf_cap.paragraphs[0]
        p_cap.text = caption
        p_cap.font.size = Pt(12)
        p_cap.font.italic = True
        p_cap.font.color.rgb = RGBColor(0x61, 0x61, 0x61)
        p_cap.alignment = PP_ALIGN.CENTER

    return prs


# ── Generate English PPTX ──
prs_en = create_pptx(
    figures_en, "en",
    "The Hidden Cost of Reproducibility:\nCommercial Software Dependency in Published Research"
)
out_en = OUTPUT_DIR / "figures_english.pptx"
prs_en.save(str(out_en))
print(f"English PPTX saved: {out_en}")

# ── Generate Japanese PPTX ──
prs_jp = create_pptx(
    figures_jp, "jp",
    "\u518d\u73fe\u6027\u306e\u96a0\u308c\u305f\u30b3\u30b9\u30c8\uff1a\n\u5b66\u8853\u7814\u7a76\u306b\u304a\u3051\u308b\u5546\u7528\u30bd\u30d5\u30c8\u30a6\u30a7\u30a2\u4f9d\u5b58\u3068\u30d0\u30fc\u30b8\u30e7\u30f3\u30a2\u30af\u30bb\u30b7\u30d3\u30ea\u30c6\u30a3\u30ae\u30e3\u30c3\u30d7"
)
out_jp = OUTPUT_DIR / "figures_japanese.pptx"
prs_jp.save(str(out_jp))
print(f"Japanese PPTX saved: {out_jp}")

print("\nBoth PPTX files created successfully!")
