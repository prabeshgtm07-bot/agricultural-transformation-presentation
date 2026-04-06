"""
Nepal Agricultural Transformation Roadmap - PowerPoint Generator
"""

import json
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

# Colors
COLOR_PRIMARY_BLUE = RGBColor(0, 51, 102)
COLOR_FOREST_GREEN = RGBColor(45, 80, 22)
COLOR_BURNT_ORANGE = RGBColor(217, 119, 6)
COLOR_LIGHT_GRAY = RGBColor(243, 244, 246)
COLOR_DARK_GRAY = RGBColor(31, 41, 55)
COLOR_WHITE = RGBColor(255, 255, 255)

# Load content
with open('data/content.json', 'r') as f:
    content = json.load(f)

# Create presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

def add_title_slide(title, subtitle, footer_left, footer_right):
    """Add title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_PRIMARY_BLUE
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.alignment = PP_ALIGN.CENTER
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(2))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    p = subtitle_frame.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(32)
    p.font.color.rgb = COLOR_BURNT_ORANGE
    p.alignment = PP_ALIGN.CENTER
    
    footer_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(4), Inches(0.6))
    footer_frame = footer_box.text_frame
    p = footer_frame.paragraphs[0]
    p.text = footer_left
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_WHITE
    
    footer_box2 = slide.shapes.add_textbox(Inches(5.5), Inches(6.8), Inches(4), Inches(0.6))
    footer_frame2 = footer_box2.text_frame
    p = footer_frame2.paragraphs[0]
    p.text = footer_right
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_WHITE
    p.alignment = PP_ALIGN.RIGHT

def add_content_slide(title, bullets, color_scheme='blue'):
    """Add standard content slide with bullets"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_WHITE
    
    title_shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.8))
    title_shape.fill.solid()
    if color_scheme == 'blue':
        title_shape.fill.fore_color.rgb = COLOR_PRIMARY_BLUE
    elif color_scheme == 'green':
        title_shape.fill.fore_color.rgb = COLOR_FOREST_GREEN
    elif color_scheme == 'orange':
        title_shape.fill.fore_color.rgb = COLOR_BURNT_ORANGE
    title_shape.line.color.rgb = title_shape.fill.fore_color.rgb
    
    title_frame = title_shape.text_frame
    title_frame.word_wrap = True
    title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    
    left = Inches(0.75)
    top = Inches(1.2)
    width = Inches(8.5)
    height = Inches(5.5)
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    
    for i, bullet in enumerate(bullets):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        
        p.text = bullet
        p.level = 0
        p.font.size = Pt(18)
        p.font.color.rgb = COLOR_DARK_GRAY
        p.space_before = Pt(6)
        p.space_after = Pt(6)
    
    footer_shape = slide.shapes.add_textbox(Inches(9.2), Inches(7.1), Inches(0.6), Inches(0.3))
    footer_frame = footer_shape.text_frame
    p = footer_frame.paragraphs[0]
    p.text = f"{len(prs.slides)}"
    p.font.size = Pt(10)
    p.font.color.rgb = COLOR_LIGHT_GRAY
    p.alignment = PP_ALIGN.RIGHT

# SLIDES

# SLIDE 1: Title
add_title_slide(
    "Nepal Agricultural Transformation",
    "Roadmap 2025–2030",
    "Ministry of Agriculture & Livestock Development",
    "Government of Nepal | April 2026"
)

# SLIDE 2: Crisis
add_content_slide(
    "The Crisis in One Sentence",
    [
        "Nepal operates at 50% of agricultural potential",
        "3.4 million farmers trapped in subsistence",
        "730,000 workers leave annually to migration",
        "NPR 28 billion subsidy reaches only 30-40%",
        "Average farm: 0.6 hectares",
        "Only 36% irrigation coverage"
    ],
    'orange'
)

# SLIDE 3
add_content_slide(
    "Operating at Half Potential",
    [
        "3.4 million farm households",
        "Average holding: 0.6 hectares",
        "Paddy yield: 3.1 tonnes/ha (vs 5.5+ potential)",
        "Irrigation: 36% coverage",
        "Agricultural growth: ~1% annually",
        "Labour exodus: 730,000 workers/year"
    ],
    'blue'
)

# SLIDE 4
add_content_slide(
    "The Budget Paradox",
    [
        "Total agriculture budget: NPR 57.5 Bn/year",
        "Fertilizer subsidies: NPR 27.95 Bn (49%)",
        "Reaches only 30-40% of farmers",
        "Elite capture rampant",
        "Budget growth flatlined at 0.34% YoY",
        "Money wasted, not reaching poorest"
    ],
    'orange'
)

# SLIDE 5
add_content_slide(
    "The Income Trap – Migration is Rational",
    [
        "Subsistence farm income: NPR 30-50K/year",
        "International wage labor: NPR 150-300K/year",
        "Income gap: 3–10× higher abroad",
        "Remittances: USD 8 billion annually",
        "Families don't farm if money arrives monthly",
        "Migration continues until farm income rises"
    ],
    'blue'
)

# SLIDE 6
add_content_slide(
    "10 Root Causes Creating Vicious Cycles",
    [
        "1. Structural land fragmentation (0.6 ha)",
        "2. Irrigation crisis (36% coverage)",
        "3. Input deficiency (67 kg/ha vs 164 in India)",
        "4. Climate change acceleration",
        "5. Labour exodus (3–10× wage gap)",
        "6. Market failure (farmers get 30-40% retail)",
        "7. Finance starvation (94% excluded)",
        "8. Human capital collapse (1:1500 ratio)",
        "9. Soil degradation",
        "10. Governance failure"
    ],
    'blue'
)

# SLIDE 7
add_content_slide(
    "Vicious Cycle 1: Productivity Trap",
    [
        "Small holding + low inputs → Low productivity",
        "Low productivity → Low income",
        "Low income → Out-migration",
        "Migration → Labour shortage → Fallow land",
        "Fallow land → Reduced output",
        "RESULT: Farming cannot intensify"
    ],
    'orange'
)

# SLIDE 8
add_content_slide(
    "Vicious Cycle 2: Subsistence Trap",
    [
        "No collateral → No credit access",
        "No credit → Cannot invest",
        "No investment → Subsistence productivity",
        "Tiny surplus → Middlemen control",
        "Poor prices → Cannot reinvest",
        "RESULT: Trapped in subsistence"
    ],
    'green'
)

# SLIDE 9
add_content_slide(
    "Vicious Cycle 3: Climate Vulnerability",
    [
        "Limited irrigation → Monsoon farming",
        "Erratic monsoons → Cannot plan",
        "Cannot plan → No adaptation possible",
        "Adaptation needs capital → Credit unavailable",
        "No resilience → Vulnerable to shocks",
        "RESULT: Locked in high-risk agriculture"
    ],
    'blue'
)

# SLIDE 10
add_content_slide(
    "Vicious Cycle 4: Soil Collapse",
    [
        "Low fertilizer + no organic input → Nutrient mining",
        "Hill erosion: 10-15 tonnes/ha annually",
        "Declining fertility → Yield falls",
        "Falling yields → Lower income",
        "Lower income → Cannot afford conservation",
        "RESULT: Self-accelerating degradation"
    ],
    'orange'
)

# SLIDE 11
add_content_slide(
    "Why Current Approaches Fail",
    [
        "Subsidy without irrigation = wasted inputs",
        "Training without credit = cannot implement",
        "Seeds without extension = planted wrong",
        "Cooperatives without consolidation = no scale",
        "Market access without infrastructure = exploited",
        "CORE PROBLEM: Treating symptoms, not system"
    ],
    'orange'
)

# SLIDE 12
add_content_slide(
    "Governance Paralysis",
    [
        "Agriculture scattered: Federal, 7 Provinces, 753 Local",
        "Federalisation confusion → No clear mandates",
        "200+ fragmented programs → No strategy",
        "Budget execution only 70%",
        "Corruption & elite capture rampant",
        "Research never reaches farms"
    ],
    'blue'
)

# SLIDE 13
add_content_slide(
    "The Subsidy Trap",
    [
        "Blanket design → Elite capture disproportionate",
        "Poor farmers get only 30-40%",
        "Distorts markets → Suppresses investment",
        "Zero lasting capacity",
        "Crowds out capital investment",
        "Like feeding someone without teaching them"
    ],
    'orange'
)

# SLIDE 14
add_content_slide(
    "8 Reform Pillars (SOLUTIONS)",
    [
        "1. Land Consolidation – 100,000 ha by 2030",
        "2. Irrigation – 36% to 65% coverage",
        "3. Inputs – 67 to 140 kg/ha fertilizer",
        "4. Markets – Farmer share 30% to 50-60%",
        "5. Finance – Credit 6% to 35%",
        "6. Extension – 1:1500 to 1:500 ratio",
        "7. Climate & Soil – Adaptive varieties",
        "8. Governance – Clarity & accountability"
    ],
    'green'
)

# SLIDE 15
add_content_slide(
    "Pillar 1: Land Consolidation",
    [
        "PROBLEM: 0.6 ha average; 50% undocumented",
        "Voluntary pooling – NPR 80K/ha grants",
        "Land documentation – Digitize all holdings",
        "Family Land Trust Law – Inherit shares",
        "TARGET 2030: 100,000 ha pooled; 90% documented",
        "INVESTMENT: NPR 18 Bn"
    ],
    'green'
)

# SLIDE 16
add_content_slide(
    "Pillar 2: Irrigation Expansion",
    [
        "PROBLEM: 36% coverage; 80% rain in 4 months",
        "Small-scale schemes – 500,000 new hectares",
        "Water harvesting – 200,000 ponds; 500 check dams",
        "Groundwater governance – Clear authority",
        "TARGET 2030: 65% coverage; year-round farming",
        "INVESTMENT: NPR 85 Bn | HIGHEST ROI"
    ],
    'green'
)

# SLIDE 17
add_content_slide(
    "Pillar 3: Input System Overhaul",
    [
        "PROBLEM: 67 kg/ha fertilizer vs 164 in India",
        "Smart Farmer Card – Quantity-capped subsidy",
        "Subsidy reform – NPR 28 Bn to NPR 14 Bn",
        "Seed production – 7 provincial farms",
        "Fertilizer security – Multi-source suppliers",
        "TARGET 2030: 140 kg/ha; 90%+ reach; 60% yield gain"
    ],
    'green'
)

# SLIDE 18
add_content_slide(
    "Pillar 4: Market Transformation",
    [
        "PROBLEM: Farmers 30-40% of retail; 30-40% losses",
        "77 District Aggregation Centres – Cold storage",
        "7 Provincial Processing Hubs – Value-add",
        "Digital Trade Platform – Direct to buyers",
        "'Himalayan Origin' brand – Premium",
        "TARGET 2030: Share 50-60%; losses 10-15%; exports 3×"
    ],
    'green'
)

# SLIDE 19
add_content_slide(
    "Pillar 5: Agricultural Finance",
    [
        "PROBLEM: Interest 18-24%; 6.7% pre-harvest credit",
        "ACGF – NPR 20 Bn government capital",
        "70% guarantee → Interest drops to 10-12%",
        "Unlocks NPR 100 Bn private lending (5:1)",
        "Harvest-linked repayment; warehouse receipts",
        "TARGET 2030: 35% with formal credit"
    ],
    'green'
)

# SLIDE 20
add_content_slide(
    "Pillar 6: Extension & Human Capital",
    [
        "PROBLEM: 1 technician per 1,500 farmers",
        "Hire 6,000 JAEOs – Home-district deployment",
        "Curriculum reform – Climate-smart mandatory",
        "Farmer Field Schools – 2 per JAEO/year",
        "NAKP app – Pest diagnosis; prices; weather",
        "TARGET 2030: 1:500 ratio; 300K farmers/year trained"
    ],
    'green'
)

# SLIDE 21
add_content_slide(
    "Pillar 7: Climate & Soil Health",
    [
        "PROBLEM: Erratic monsoons; declining fertility",
        "Climate vulnerability mapping – 77 districts",
        "Adaptive varieties – Flood/drought tolerant",
        "Soil Health Cards – Free testing",
        "Organic matter recovery – Compost subsidies",
        "TARGET 2030: 50%+ climate-smart; soil OM doubled"
    ],
    'green'
)

# SLIDE 22
add_content_slide(
    "Pillar 8: Governance Reform",
    [
        "PROBLEM: Fragmented; 200+ programs; 70% execution",
        "Functions Clarity Act – Define each tier",
        "National Coordination Council – Quarterly meetings",
        "8 Flagship Programmes – Consolidate 200 items",
        "Quarterly Public Dashboard – Transparent",
        "TARGET 2030: 85%+ execution; zero leakage"
    ],
    'green'
)

# SLIDE 23
add_content_slide(
    "Financial Logic: NPR 340 Bn (Self-Financed)",
    [
        "Current trajectory: NPR 287.5 Bn",
        "Proposed: NPR 340 Bn",
        "Incremental: NPR 52.5 Bn (18% increase)",
        "",
        "FUNDING: Subsidy reform (70 Bn) + loans + leverage",
        "STATUS: FULLY FUNDED"
    ],
    'blue'
)

# SLIDE 24
add_content_slide(
    "Five-Year Timeline",
    [
        "YEAR 1 (NPR 55 Bn): Foundation – Laws, pilots",
        "YEAR 2 (NPR 81 Bn): Build – 200K ha irrigation",
        "YEAR 3 (NPR 87 Bn): Peak – 320K ha; scaling",
        "YEAR 4 (NPR 75 Bn): Consolidate – 430K ha total",
        "YEAR 5 (NPR 60 Bn): Sustain – 65% coverage"
    ],
    'blue'
)

# SLIDE 25
add_content_slide(
    "2030 Transformation Targets",
    [
        "Irrigated land: 36% → 65% (+81%)",
        "Paddy yield: 3.1 → 5.5+ tonnes/ha (1.77×)",
        "Fertilizer: 67 → 140+ kg/ha (2.08×)",
        "Formal credit: 6% → 35%+ (5×)",
        "Post-harvest losses: 30-40% → 10-15%",
        "Extension ratio: 1:1,500 → 1:500 (3×)",
        "Agricultural growth: 1% → 4-5% annually"
    ],
    'green'
)

# SLIDE 26
add_content_slide(
    "The Choice: Inaction vs Transformation",
    [
        "STATUS QUO:",
        "• Subsidy 30-40% reach; irrigation 36%",
        "• Labour exodus 730K/year; 32% fallow",
        "• Income stagnant; 94% credit excluded",
        "• Growth ~1%; food insecurity rising",
        "",
        "TRANSFORMATION:",
        "• Subsidy 90%+ reach; irrigation 65%",
        "• Exodus slows; farm income competitive",
        "• 35%+ formal credit; mechanization viable",
        "• Growth 4-5%; food self-sufficient"
    ],
    'blue'
)

# SLIDE 27
add_content_slide(
    "Why This Roadmap Works",
    [
        "POLITICAL: Targets elite capture, not poor",
        "FINANCIAL: Self-financed through reallocation",
        "TECHNICAL: All components proven elsewhere",
        "INSTITUTIONAL: Clear accountability",
        "  • Quarterly dashboards create pressure",
        "  • Farmer feedback keeps responsive",
        "  • Independent evaluation (Year 3)",
        "  • Laws bind future governments"
    ],
    'green'
)

# SLIDE 28
add_content_slide(
    "100-Day Quick Wins",
    [
        "DAYS 1-30: Fertilizer MOUs; NRB directive",
        "  • Unlock NPR 15-20 Bn private lending",
        "",
        "DAYS 1-15: Announce Land Pooling",
        "",
        "DAYS 30-60: Climate maps; recruit 500 JAEOs",
        "",
        "DAYS 60-100: Smart Card pilots; sign loans",
        "",
        "RESULT: Momentum; 500 JAEOs deployed"
    ],
    'orange'
)

# SLIDE 29
add_content_slide(
    "Three Core Truths",
    [
        "1. STRUCTURAL PROBLEM = STRUCTURAL SOLUTION",
        "   All 8 pillars activate simultaneously",
        "",
        "2. REALLOCATION, NOT BANKRUPTCY",
        "   NPR 28 Bn → NPR 14 Bn smart + NPR 14 Bn capital",
        "",
        "3. POLITICAL WILL IS THE CONSTRAINT",
        "   Technology works. Money available.",
        "",
        "DO WE SUBSIDIZE POVERTY OR INVEST IN PROSPERITY?"
    ],
    'blue'
)

# Save
output_dir = 'output'
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, 'Nepal_Agricultural_Transformation_Roadmap.pptx')
prs.save(output_path)
print(f"✓ PowerPoint created: {output_path}")
print(f"✓ Total slides: {len(prs.slides)}")
