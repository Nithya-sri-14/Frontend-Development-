from pathlib import Path
from openpyxl import Workbook
import re

repo = Path('/Users/user/Frontend-Development-')
assets_root = repo / 'assets' / 'expected-outputs'
assets_root.mkdir(parents=True, exist_ok=True)

owner = 'Nithya-sri-14'
repo_name = 'Frontend-Development-'
branch = 'main'

language = 'Frontend Development'
levels = ['Beginner', 'Intermediate', 'Hard']
topics = [
    'HTML',
    'Semantic HTML',
    'Forms',
    'Tables',
    'CSS',
    'Box Model',
    'Flexbox',
    'Grid',
    'Media Queries',
    'Responsive Design',
]

subtopics = {
    'HTML': ['Document Structure', 'Head & Meta', 'Links & Media', 'Lists & Grouping', 'Sectioning'],
    'Semantic HTML': ['Landmarks', 'Accessible Structure', 'Figure & Caption', 'Time & Address', 'Content Hierarchy'],
    'Forms': ['Input Types', 'Labels', 'Validation', 'Fieldsets', 'Form UX'],
    'Tables': ['Basic Table', 'Spans', 'Header Scope', 'Captions', 'Responsive Table Patterns'],
    'CSS': ['Selectors', 'Typography', 'Colors', 'Pseudo Classes', 'Variables'],
    'Box Model': ['Margin', 'Padding', 'Border', 'Sizing', 'Spacing Systems'],
    'Flexbox': ['Axis Alignment', 'Wrapping', 'Order', 'Grow/Shrink', 'Component Layout'],
    'Grid': ['Tracks', 'Areas', 'Auto Placement', 'Gaps', 'Dashboard Patterns'],
    'Media Queries': ['Breakpoints', 'Mobile-first', 'Typography Scale', 'Layout Switch', 'Interaction Changes'],
    'Responsive Design': ['Fluid Layout', 'Fluid Media', 'Navigation Patterns', 'Card Systems', 'Adaptive Components'],
}

level_verbs = {
    'Beginner': 'Build',
    'Intermediate': 'Implement',
    'Hard': 'Engineer',
}

level_constraints = {
    'Beginner': [
        'Use only HTML and CSS',
        'Do not use JavaScript',
        'Keep code beginner-friendly and readable',
        'Use semantic tags where applicable',
    ],
    'Intermediate': [
        'Use clean structure and reusable classes',
        'Support keyboard focus states',
        'Avoid framework dependencies',
        'Write mobile-first CSS rules first',
    ],
    'Hard': [
        'Optimize for accessibility and maintainability',
        'Use scalable architecture with reusable patterns',
        'Avoid layout shifts between breakpoints',
        'Keep specificity low and predictable',
    ],
}

def slugify(text: str) -> str:
    return re.sub(r'[^a-z0-9]+', '-', text.lower()).strip('-')

def build_title(level: str, topic: str, subtopic: str, idx: int) -> str:
    verb = level_verbs[level]
    return f"{verb} {topic} Practice {idx:02d}: {subtopic}"

def scenario(level: str, topic: str, subtopic: str, idx: int) -> str:
    if level == 'Beginner':
        return f"You are creating a starter {topic.lower()} exercise for students in an LMS lab. The target is a clear, simple output focused on {subtopic.lower()}."
    if level == 'Intermediate':
        return f"You are building a production-like {topic.lower()} challenge module where students must apply {subtopic.lower()} with structure, consistency, and accessibility checks."
    return f"You are designing an advanced frontend assessment where students must solve a realistic {topic.lower()} implementation centered on {subtopic.lower()} under strict quality constraints."

def instruction_hint(level: str, topic: str, subtopic: str) -> str:
    if level == 'Beginner':
        return f"Start with a minimal semantic HTML skeleton, then add CSS classes step by step to match the expected {topic.lower()} layout."
    if level == 'Intermediate':
        return f"Plan component structure first, then implement reusable class patterns and test how {subtopic.lower()} behaves at multiple viewport widths."
    return f"Define a scalable styling strategy before coding, then validate accessibility, responsiveness, and edge states for the {subtopic.lower()} requirement."

def description(level: str, topic: str, subtopic: str, idx: int) -> str:
    if level == 'Beginner':
        return f"Create a beginner task for {topic} using the subtopic {subtopic}. Build a single-page output that matches the reference image exactly for challenge #{idx}."
    if level == 'Intermediate':
        return f"Create an intermediate task for {topic} using {subtopic}. Students must produce a reusable, responsive section that matches the reference output for challenge #{idx}."
    return f"Create a hard-level task for {topic} using {subtopic}. Students must engineer a robust, accessible UI output with production-grade structure for challenge #{idx}."

def constraint_text(level: str, idx: int) -> str:
    rules = level_constraints[level]
    pick = rules[(idx - 1) % len(rules)]
    return f"{pick}; submission must render correctly at required viewport sizes."


def write_svg(path: Path, level: str, topic: str, title: str, idx: int):
    colors = {
        'Beginner': ('#0f766e', '#ccfbf1'),
        'Intermediate': ('#1d4ed8', '#dbeafe'),
        'Hard': ('#b45309', '#fef3c7'),
    }
    stroke, bg = colors[level]
    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="630" viewBox="0 0 1200 630" role="img" aria-label="Expected output preview">
  <rect width="1200" height="630" fill="{bg}"/>
  <rect x="40" y="40" width="1120" height="550" rx="20" fill="white" stroke="{stroke}" stroke-width="6"/>
  <text x="80" y="120" font-family="Arial, sans-serif" font-size="36" font-weight="700" fill="{stroke}">Expected Output</text>
  <text x="80" y="175" font-family="Arial, sans-serif" font-size="28" font-weight="600" fill="#111827">{level} • {topic} • Q{idx:02d}</text>
  <foreignObject x="80" y="210" width="1040" height="330">
    <div xmlns="http://www.w3.org/1999/xhtml" style="font-family: Arial, sans-serif; font-size: 24px; color: #1f2937; line-height: 1.4;">
      <p style="margin:0;">{title}</p>
      <p style="margin-top:20px;">Replicate this layout and styling in your HTML/CSS submission exactly as shown in this reference.</p>
    </div>
  </foreignObject>
  <rect x="80" y="500" width="280" height="56" rx="12" fill="{stroke}"/>
  <text x="110" y="536" font-family="Arial, sans-serif" font-size="24" fill="white">Target: Visual Match</text>
</svg>'''
    path.write_text(svg, encoding='utf-8')

rows = []
for level in levels:
    for topic in topics:
        subs = subtopics[topic]
        for i in range(1, 21):
            sub = subs[(i - 1) % len(subs)]
            title = build_title(level, topic, sub, i)
            desc = description(level, topic, sub, i)
            scen = scenario(level, topic, sub, i)
            hint = instruction_hint(level, topic, sub)
            cons = constraint_text(level, i)

            topic_slug = slugify(topic)
            level_slug = slugify(level)
            file_name = f"q{i:02d}-{slugify(sub)}.svg"
            svg_path = assets_root / level_slug / topic_slug / file_name
            svg_path.parent.mkdir(parents=True, exist_ok=True)
            write_svg(svg_path, level, topic, title, i)

            rel = svg_path.relative_to(repo).as_posix()
            example_link = f"https://raw.githubusercontent.com/{owner}/{repo_name}/{branch}/{rel}"

            rows.append([
                language,
                level,
                topic,
                title,
                desc,
                scen,
                hint,
                cons,
                example_link,
            ])

wb = Workbook()
ws = wb.active
ws.title = 'Frontend Practice Bank'
headers = [
    'Language',
    'Level',
    'Topic',
    'Title',
    'Description',
    'Explanation scenario',
    'Instruction hint',
    'Constraints',
    'Example output',
]
ws.append(headers)
for row in rows:
    ws.append(row)

for col in ['A','B','C','D','E','F','G','H','I']:
    ws.column_dimensions[col].width = 35

out = repo / 'Frontend_Development_Practice_Bank.xlsx'
wb.save(out)

print(f"Created {len(rows)} questions")
print(f"Excel: {out}")
