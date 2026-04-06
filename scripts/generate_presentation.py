"""
Nepal Agricultural Transformation Roadmap - PowerPoint Generator
Complete working version with 16 slides
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

def add_title_slide(title, subtitle):
    """Add title slide - FIXED: Only 2 parameters"""
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
    
    # Footer
    footer_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.6))
    footer_frame = footer_box.text_frame
    p = footer_frame.paragraphs[0]
    p.text = "Government of Nepal | Ministry of Agriculture & Livestock Development"
    p.font.size = Pt(10)
    p.font.color.rgb = COLOR_WHITE

def add_content_slide(title, bullets, color='blue'):
    """Add standard content slide"""
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

# SLIDE 1: Title
add_title_slide(
    "Nepal Agricultural Transformation",
    "Roadmap 2025–2030"
)

# SLIDE 2: Crisis
add_content_slide("The Crisis in One Sentence", [
    "Nepal operates at 50% agricultural potential",
    "3.4 million farmers trapped in subsistence",
    "730,000 workers leave annually",
    "Average farm: 0.6 hectares (below viability)",
    "Only 36% irrigation coverage"
], 'orange')

# SLIDE 3: Root Causes
add_content_slide("10 Root Causes", [
    "Land fragmentation (0.6 ha average)",
    "Irrigation crisis (36% coverage)",
    "Input deficiency (67 vs 164 kg/ha in India)",
    "Labour exodus (3-10x wage gap)",
    "Market failure (30-40% farmer share)",
    "Finance starvation (94% excluded)",
    "Climate vulnerability",
    "Soil degradation (10-15 tonnes/ha erosion)",
    "Governance failure (fragmented)",
    "Human capital collapse"
], 'blue')

# SLIDE 4: Vicious Cycles
add_content_slide("4 Vicious Cycles", [
    "Cycle 1: Small holdings → Low productivity → Migration",
    "Cycle 2: No collateral → No credit → Cannot invest",
    "Cycle 3: Limited irrigation → Cannot plan → No adaptation",
    "Cycle 4: Nutrient mining → Soil collapse → Yield decline"
], 'blue')

# SLIDE 5: Why Current Approaches Fail
add_content_slide("Why Subsidies Alone Fail", [
    "Subsidy without irrigation = wasted inputs",
    "Training without credit = cannot implement",
    "Seeds without extension = planted incorrectly",
    "Cooperatives without scale = no impact",
    "All these are disconnected fixes for a system problem"
], 'orange')

# SLIDE 6: 8 Reform Pillars
add_content_slide("8 Interconnected Pillars (SOLUTIONS)", [
    "1. Land Consolidation (100K ha by 2030)",
    "2. Irrigation Expansion (36% → 65%)",
    "3. Input System (67 → 140 kg/ha)",
    "4. Markets (30% → 50-60% farmer share)",
    "5. Agricultural Finance (6% → 35% credit)",
    "6. Extension (1:1500 → 1:500 ratio)",
    "7. Climate & Soil (adaptive varieties)",
    "8. Governance (clarity & accountability)"
], 'green')

# SLIDE 7: Pillar 1 - Land
add_content_slide("Pillar 1: Land Consolidation", [
    "Problem: 0.6 ha average, 50% undocumented",
    "Solution: Voluntary pooling (NPR 80K/ha grants)",
    "Land documentation (digitize all holdings)",
    "Family Land Trust Law (inherit shares, not land)",
    "Target 2030: 100,000 ha pooled, 90% documented"
], 'green')

# SLIDE 8: Pillar 2 - Irrigation
add_content_slide("Pillar 2: Irrigation Expansion", [
    "Problem: 36% coverage, 80% rain in 4 months",
    "Solution: Small-scale schemes (500K new ha)",
    "Water harvesting (200K farm ponds)",
    "Groundwater governance",
    "Target 2030: 65% coverage, year-round farming"
], 'green')

# SLIDE 9: Pillar 4 - Markets
add_content_slide("Pillar 4: Market Transformation", [
    "Problem: Farmers get 30-40% retail, 30-40% losses",
    "Solution: 77 District Aggregation Centres (cold storage)",
    "7 Provincial Processing Hubs (value-add)",
    "Digital Trade Platform (direct buyers)",
    "Target 2030: 50-60% share, 3x exports"
], 'green')

# SLIDE 10: Pillar 5 - Finance
add_content_slide("Pillar 5: Agricultural Finance", [
    "Problem: 18-24% interest, only 6% get credit",
    "Solution: ACGF (NPR 20 Bn government capital)",
    "70% guarantee → interest drops to 10-12%",
    "Unlocks NPR 100 Bn private lending (5:1 leverage)",
    "Target 2030: 35% with formal credit"
], 'green')

# SLIDE 11: Financial Framework
add_content_slide("Financial Logic: NPR 340 Bn", [
    "Current trajectory: NPR 287.5 Bn",
    "Proposed: NPR 340 Bn (18% increase)",
    "Funding: Subsidy reform (70 Bn) + loans + leverage",
    "ACGF private leverage: NPR 100 Bn",
    "Status: FULLY FUNDED"
], 'blue')

# SLIDE 12: Timeline
add_content_slide("5-Year Implementation", [
    "Year 1 (NPR 55 Bn): Foundation - Laws, pilots",
    "Year 2 (NPR 81 Bn): Build - 200K ha irrigation added",
    "Year 3 (NPR 87 Bn): Peak - 320K ha total, scaling",
    "Year 4 (NPR 75 Bn): Consolidate - 430K ha total",
    "Year 5 (NPR 60 Bn): Sustain - 65% coverage achieved"
], 'blue')

# SLIDE 13: 2030 Targets
add_content_slide("2030 Transformation Targets", [
    "Irrigation: 36% → 65% coverage",
    "Paddy yield: 3.1 → 5.5+ tonnes/ha (1.77x)",
    "Fertilizer: 67 → 140+ kg/ha (2.08x)",
    "Formal credit: 6% → 35% (5x)",
    "Agricultural growth: 1% → 4-5% annually",
    "Farmer income: 30-50K → 150-250K NPR/year"
], 'green')

# SLIDE 14: Status Quo vs Transformation
add_content_slide("The Choice: Inaction vs Transformation", [
    "STATUS QUO: Subsidy 30-40% reach, exodus 730K/year",
    "TRANSFORMATION: Subsidy 90%+ reach, exodus slows",
    "",
    "STATUS QUO: 94% excluded from credit",
    "TRANSFORMATION: 35% with viable credit",
    "",
    "STATUS QUO: Growth ~1%, food insecurity rising",
    "TRANSFORMATION: Growth 4-5%, self-sufficient"
], 'blue')

# SLIDE 15: Why This Works
add_content_slide("Why This Roadmap Is Implementable", [
    "Political: Targets elite capture, not poor farmers",
    "Financial: Self-financed through reallocation + leverage",
    "Technical: All components proven in other countries",
    "Institutional: Clear accountability, laws bind governments",
    "Quarterly dashboards + independent evaluation"
], 'green')

# SLIDE 16: Core Truth
add_content_slide("The Closing Question", [
    "",
    "Do we continue to subsidize poverty?",
    "",
    "Or do we invest in prosperity?",
    "",
    "This roadmap is the blueprint. Political will is the constraint."
], 'blue')

# Save
output_dir = 'output'
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, 'Nepal_Agricultural_Transformation_Roadmap.pptx')
prs.save(output_path)
print(f"✓ PowerPoint created: {output_path}")
print(f"✓ Total slides: {len(prs.slides)}")
