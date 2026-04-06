"""
Nepal Agricultural Transformation Roadmap - PowerPoint Generator
Simplified version without external dependencies
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

# Colors
COLOR_PRIMARY_BLUE = RGBColor(0, 51, 102)
COLOR_FOREST_GREEN = RGBColor(45, 80, 22)
COLOR_BURNT_ORANGE = RGBColor(217, 119, 6)
COLOR_DARK_GRAY = RGBColor(31, 41, 55)
COLOR_WHITE = RGBColor(255, 255, 255)

# Create presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

def add_title_slide(title, subtitle, footer_left, footer_right):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_PRIMARY_BLUE
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.alignment = PP_ALIGN.CENTER
    
    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(2))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(32)
    p.font.color.rgb = COLOR_BURNT_ORANGE
    p.alignment = PP_ALIGN.CENTER

def add_content_slide(title, bullets, color='blue'):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title bar
    title_shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.8))
    title_shape.fill.solid()
    if color == 'blue':
        title_shape.fill.fore_color.rgb = COLOR_PRIMARY_BLUE
    elif color == 'green':
        title_shape.fill.fore_color.rgb = COLOR_FOREST_GREEN
    elif color == 'orange':
        title_shape.fill.fore_color.rgb = COLOR_BURNT_ORANGE
    
    title_frame = title_shape.text_frame
    title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    
    # Content
    text_box = slide.shapes.add_textbox(Inches(0.75), Inches(1.2), Inches(8.5), Inches(5.5))
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    
    for i, bullet in enumerate(bullets):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        p.text = bullet
        p.font.size = Pt(16)
        p.font.color.rgb = COLOR_DARK_GRAY

# SLIDES
add_title_slide(
    "Nepal Agricultural Transformation",
    "Roadmap 2025–2030",
    "Ministry of Agriculture & Livestock Development"
)

add_content_slide("The Crisis", [
    "Nepal operates at 50% agricultural potential",
    "3.4 million farmers trapped in subsistence",
    "730,000 workers leave annually",
    "Average farm: 0.6 hectares",
    "Only 36% irrigation coverage"
], 'orange')

add_content_slide("Root Causes", [
    "Land fragmentation (0.6 ha average)",
    "Irrigation crisis (36% coverage)",
    "Input deficiency (67 kg/ha vs 164 in India)",
    "Labour exodus (3–10× wage gap)",
    "Market failure (farmers get 30-40% retail)"
], 'blue')

add_content_slide("Vicious Cycles", [
    "Cycle 1: Productivity trap → Out-migration",
    "Cycle 2: No credit → Cannot invest",
    "Cycle 3: Limited water → Cannot adapt",
    "Cycle 4: Soil collapse → Yield decline",
    "All cycles interconnected and reinforcing"
], 'blue')

add_content_slide("8 Reform Pillars", [
    "1. Land Consolidation – 100,000 ha by 2030",
    "2. Irrigation Expansion – 36% to 65%",
    "3. Input System – Fertilizer 67 to 140 kg/ha",
    "4. Markets – Farmer share 30% to 50-60%",
    "5. Agricultural Finance – 6% to 35% credit",
    "6. Extension – 1:500 ratio (vs 1:1500 now)",
    "7. Climate & Soil – Adaptive varieties",
    "8. Governance – Functions clarity"
], 'green')

add_content_slide("Pillar 1: Land", [
    "Problem: 0.6 ha average; 50% undocumented",
    "Solution: Voluntary pooling (NPR 80K/ha)",
    "Land documentation – Digitize all holdings",
    "Family Land Trust Law",
    "Target 2030: 100,000 ha pooled"
], 'green')

add_content_slide("Pillar 2: Irrigation", [
    "Problem: 36% coverage; 80% rain in 4 months",
    "Solution: Small-scale schemes (500K ha new)",
    "Water harvesting (200K ponds)",
    "Groundwater governance",
    "Target 2030: 65% coverage"
], 'green')

add_content_slide("Pillar 3: Inputs", [
    "Problem: 67 kg/ha vs 164 in India",
    "Solution: Smart Farmer Card (capped subsidy)",
    "Subsidy reform: NPR 28 Bn to 14 Bn",
    "Seed production (7 provincial farms)",
    "Target 2030: 140 kg/ha"
], 'green')

add_content_slide("Pillar 4: Markets", [
    "Problem: Farmers 30-40% retail; 30-40% losses",
    "Solution: 77 Aggregation Centres (cold storage)",
    "7 Processing Hubs (value-add)",
    "Digital Trade Platform",
    "Target 2030: 50-60% share; 3× exports"
], 'green')

add_content_slide("Pillar 5: Finance", [
    "Problem: Interest 18-24%; 6% credit access",
    "Solution: ACGF (NPR 20 Bn government)",
    "70% guarantee → Interest drops to 10-12%",
    "Unlocks NPR 100 Bn private lending",
    "Target 2030: 35% formal credit"
], 'green')

add_content_slide("Pillar 6: Extension", [
    "Problem: 1 technician per 1,500 farmers",
    "Solution: Hire 6,000 JAEOs (home-district)",
    "Climate-smart curriculum",
    "Farmer Field Schools",
    "Target 2030: 1:500 ratio"
], 'green')

add_content_slide("Pillar 7: Climate & Soil", [
    "Problem: Monsoons erratic; soil declining",
    "Solution: Climate vulnerability mapping",
    "Adaptive crop varieties",
    "Soil Health Cards (free testing)",
    "Target 2030: 50%+ climate-smart adoption"
], 'green')

add_content_slide("Pillar 8: Governance", [
    "Problem: Fragmented; 200+ programs; 70% execution",
    "Solution: Functions Clarity Act",
    "National Coordination Council",
    "8 Flagship Programmes (consolidate)",
    "Target 2030: 85%+ execution"
], 'green')

add_content_slide("Financial Logic", [
    "Current budget: NPR 287.5 Bn (5 years)",
    "Proposed: NPR 340 Bn (18% increase)",
    "Funding: Subsidy reform + loans + leverage",
    "ACGF private leverage: NPR 100 Bn",
    "Status: FULLY FUNDED"
], 'blue')

add_content_slide("5-Year Timeline", [
    "Year 1 (NPR 55 Bn): Foundation – Laws, pilots",
    "Year 2 (NPR 81 Bn): Build – 200K ha irrigation",
    "Year 3 (NPR 87 Bn): Peak – 320K ha; scaling",
    "Year 4 (NPR 75 Bn): Consolidate – 430K ha",
    "Year 5 (NPR 60 Bn): Sustain – 65% coverage"
], 'blue')

add_content_slide("2030 Targets", [
    "Irrigation: 36% → 65% coverage",
    "Paddy yield: 3.1 → 5.5+ tonnes/ha",
    "Fertilizer: 67 → 140+ kg/ha",
    "Formal credit: 6% → 35%",
    "Agricultural growth: 1% → 4-5% annually",
    "Farmer income: 30-50K → 150-250K NPR/year"
], 'green')

add_content_slide("The Choice", [
    "STATUS QUO:",
    "Subsidy 30-40% reach; irrigation 36%",
    "Labour exodus 730K/year",
    "",
    "TRANSFORMATION:",
    "Subsidy 90%+ reach; irrigation 65%",
    "Agricultural growth 4-5%; food self-sufficient"
], 'blue')

add_content_slide("Why This Works", [
    "Political: Targets elite capture, not poor",
    "Financial: Self-financed through reallocation",
    "Technical: All components proven elsewhere",
    "Institutional: Clear accountability",
    "Laws bind future governments"
], 'green')

add_content_slide("DO WE SUBSIDIZE POVERTY", [
    "Or invest in prosperity?",
    "",
    "This roadmap provides the blueprint",
    "for structural transformation",
    "",
    "Political will is the only constraint"
], 'blue')

# Save
output_dir = 'output'
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, 'Nepal_Agricultural_Transformation_Roadmap.pptx')
prs.save(output_path)
print(f"✓ PowerPoint created: {output_path}")
print(f"✓ Total slides: {len(prs.slides)}")
