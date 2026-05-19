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

# 20 unique build-from-scratch prompts per topic
question_templates = {
    'HTML': [
        'Personal Portfolio Landing', 'Course Catalog Section', 'Blog Post Layout', 'Pricing Intro Section', 'FAQ Page Skeleton',
        'Team Profile Page', 'Event Schedule Page', 'Product Showcase Page', 'Travel Guide Article', 'Restaurant Menu Page',
        'Job Listing Page', 'Startup About Page', 'Support Center Page', 'Photo Story Page', 'Book Details Page',
        'News Highlights Page', 'Workshop Agenda Page', 'Community Announcement Page', 'Fitness Program Page', 'Podcast Episode Page'
    ],
    'Semantic HTML': [
        'Accessible News Portal', 'Podcast Episode Detail', 'Recipe Instruction Page', 'Conference Session Overview', 'Healthcare Article Page',
        'E-commerce Product Detail', 'University Department Page', 'Local Business Profile', 'Nonprofit Campaign Page', 'City Event Bulletin',
        'Learning Path Overview', 'Magazine Feature Story', 'Open Source Project Page', 'Travel Destination Guide', 'Film Review Article',
        'Museum Exhibit Page', 'Charity Impact Report', 'Startup Case Study', 'Research Summary Page', 'Technology Digest Page'
    ],
    'Forms': [
        'Student Registration Form', 'Newsletter Signup Form', 'Appointment Booking Form', 'Course Enrollment Form', 'Feedback Submission Form',
        'Support Ticket Form', 'Job Application Form', 'Hotel Reservation Form', 'Medical Intake Form', 'Event RSVP Form',
        'Profile Edit Form', 'Payment Details Form', 'Survey Response Form', 'Volunteer Signup Form', 'Internship Application Form',
        'Password Reset Form', 'Contest Entry Form', 'Shipping Address Form', 'Restaurant Reservation Form', 'Contact Us Form'
    ],
    'Tables': [
        'Student Marks Table', 'Monthly Budget Table', 'Product Inventory Table', 'Exam Timetable', 'Employee Shift Table',
        'Flight Schedule Table', 'Sales Performance Table', 'Project Milestone Table', 'Attendance Register Table', 'Order Tracking Table',
        'Subscription Plan Table', 'Hospital Duty Roster', 'Course Comparison Table', 'Quarterly Revenue Table', 'Library Catalog Table',
        'Support Response Metrics', 'Website Traffic Summary', 'Training Calendar Table', 'Sports League Standings', 'Maintenance Log Table'
    ],
    'CSS': [
        'Hero Banner Styling', 'Feature Card Styling', 'Testimonial Block Styling', 'Alert Component Styling', 'Sidebar Theme Styling',
        'Footer Styling System', 'Navigation Styling', 'Button State Styling', 'Typography Scale Styling', 'Color Token Styling',
        'Pricing Card Styling', 'Dashboard Widget Styling', 'Landing CTA Styling', 'Tag Badge Styling', 'Timeline Item Styling',
        'Profile Summary Styling', 'Search Bar Styling', 'Notification Panel Styling', 'Info Section Styling', 'Promo Block Styling'
    ],
    'Box Model': [
        'Profile Card Spacing', 'Checkout Panel Spacing', 'Feature Tile Framing', 'Form Field Spacing', 'Promo Banner Framing',
        'Article Content Widths', 'Info Block Layering', 'Review Card Breathing Space', 'Sidebar Item Padding', 'Modal Body Spacing',
        'Callout Box Boundaries', 'Avatar Card Balance', 'Stats Widget Sizing', 'Product Tile Alignment', 'List Row Spacing',
        'Media Caption Container', 'Toolbar Item Gaps', 'Card Stack Rhythm', 'Section Padding System', 'Compact Panel Sizing'
    ],
    'Flexbox': [
        'Navbar Alignment System', 'Card Row Distribution', 'Toolbar Action Layout', 'Profile Header Alignment', 'Pricing Column Alignment',
        'Form Action Row', 'Comment Thread Header', 'Dashboard Topbar', 'Mobile Menu Layout', 'Feature Split Section',
        'Rating Row Layout', 'Media Object Pattern', 'Tag Cloud Wrapping', 'Checkout Summary Row', 'Badge + Title Alignment',
        'Gallery Caption Row', 'Task Item Layout', 'Search Result Header', 'Action Button Cluster', 'Footer Link Alignment'
    ],
    'Grid': [
        'Analytics Dashboard Grid', 'Course Card Grid', 'Photo Gallery Grid', 'Product Listing Grid', 'Magazine Layout Grid',
        'Admin Panel Grid', 'Learning Modules Grid', 'KPI Metrics Grid', 'Team Directory Grid', 'News Masonry-like Grid',
        'Portfolio Showcase Grid', 'Calendar Overview Grid', 'Pricing Comparison Grid', 'Service Tiles Grid', 'Feature Matrix Grid',
        'Gallery Sidebar Grid', 'Blog Highlights Grid', 'Stats + Table Grid', 'Widget Board Grid', 'Case Studies Grid'
    ],
    'Media Queries': [
        'Responsive Navbar Switch', 'Two-to-One Column Shift', 'Typography Scale by Breakpoint', 'Card Stack Reflow', 'Table-to-Cards Switch',
        'Sidebar Collapse Pattern', 'Hero Height Adaptation', 'Image Ratio Adaptation', 'Button Group Wrap', 'Dashboard Block Rearrangement',
        'Form Layout Breakpoint', 'Footer Column Collapse', 'Gallery Column Reduction', 'Section Spacing Adaptation', 'Sticky Header Toggle',
        'Visibility by Viewport', 'Grid Area Reassignment', 'CTA Layout Shift', 'Pricing Cards Reorder', 'Mobile-first Component Tuning'
    ],
    'Responsive Design': [
        'Responsive Landing Page', 'Adaptive Course Page', 'Fluid Product Gallery', 'Multi-device Blog Layout', 'Responsive Admin Overview',
        'Mobile-first Pricing Page', 'Adaptive Event Website', 'Responsive Portfolio Site', 'Restaurant Responsive Home', 'Travel Responsive Guide',
        'Responsive Help Center', 'Adaptive SaaS Landing', 'Responsive News Homepage', 'Responsive Education Portal', 'Adaptive Community Board',
        'Responsive Profile Dashboard', 'Responsive Booking Interface', 'Adaptive Ecommerce Section', 'Responsive Magazine Layout', 'Adaptive Startup Website'
    ],
}

level_style = {
    'Beginner': {
        'bg': '#f0fdfa', 'card': '#ffffff', 'accent': '#0f766e',
        'note': 'Beginner target: clean structure, basic spacing, and clear typography.'
    },
    'Intermediate': {
        'bg': '#eff6ff', 'card': '#ffffff', 'accent': '#1d4ed8',
        'note': 'Intermediate target: reusable classes, polished states, and responsive behavior.'
    },
    'Hard': {
        'bg': '#fff7ed', 'card': '#ffffff', 'accent': '#c2410c',
        'note': 'Hard target: production-grade structure, accessibility, and exact multi-breakpoint fidelity.'
    },
}

constraints_by_level = {
    'Beginner': [
        'Use only HTML + CSS (no JS).',
        'Match spacing and font sizes exactly as shown.',
        'Use semantic elements for all sections.',
        'Keep selectors simple and readable.',
    ],
    'Intermediate': [
        'Use mobile-first CSS with at least one breakpoint.',
        'Implement visible focus states for interactive elements.',
        'Use reusable utility/component classes.',
        'Avoid fixed pixel heights unless required by design.',
    ],
    'Hard': [
        'Support at least 3 breakpoints (mobile/tablet/desktop).',
        'Preserve visual hierarchy and spacing rhythm across sizes.',
        'Keep accessibility contrast and keyboard focus compliance.',
        'Use scalable architecture with low-specificity selectors.',
    ],
}

def slugify(text: str) -> str:
    return re.sub(r'[^a-z0-9]+', '-', text.lower()).strip('-')

def render_svg(path: Path, level: str, topic: str, title: str, idx: int):
    style = level_style[level]
    accent = style['accent']
    bg = style['bg']
    card = style['card']

    # Deterministic layout variations to keep references visually distinct.
    variant = idx % 4
    if variant == 0:
        blocks = '''
  <rect x="90" y="210" width="320" height="140" rx="14" fill="#e2e8f0"/>
  <rect x="430" y="210" width="320" height="140" rx="14" fill="#dbeafe"/>
  <rect x="770" y="210" width="320" height="140" rx="14" fill="#fef3c7"/>
  <rect x="90" y="370" width="1000" height="130" rx="14" fill="#f1f5f9"/>
'''
    elif variant == 1:
        blocks = '''
  <rect x="90" y="210" width="220" height="290" rx="14" fill="#e2e8f0"/>
  <rect x="330" y="210" width="760" height="90" rx="14" fill="#dbeafe"/>
  <rect x="330" y="320" width="360" height="180" rx="14" fill="#fef3c7"/>
  <rect x="730" y="320" width="360" height="180" rx="14" fill="#ede9fe"/>
'''
    elif variant == 2:
        blocks = '''
  <rect x="90" y="210" width="1000" height="90" rx="14" fill="#dbeafe"/>
  <rect x="90" y="320" width="490" height="180" rx="14" fill="#e2e8f0"/>
  <rect x="600" y="320" width="490" height="180" rx="14" fill="#fef3c7"/>
'''
    else:
        blocks = '''
  <rect x="90" y="210" width="1000" height="140" rx="14" fill="#f1f5f9"/>
  <rect x="90" y="370" width="320" height="130" rx="14" fill="#dbeafe"/>
  <rect x="430" y="370" width="320" height="130" rx="14" fill="#e2e8f0"/>
  <rect x="770" y="370" width="320" height="130" rx="14" fill="#fef3c7"/>
'''

    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <rect width="1280" height="720" fill="{bg}"/>
  <rect x="50" y="50" width="1180" height="620" rx="22" fill="{card}" stroke="{accent}" stroke-width="5"/>
  <text x="90" y="110" font-family="Arial, sans-serif" font-size="34" font-weight="700" fill="{accent}">{title}</text>
  <text x="90" y="155" font-family="Arial, sans-serif" font-size="23" fill="#0f172a">{level} • {topic} • UI Reference #{idx:02d}</text>
  {blocks}
  <rect x="90" y="535" width="460" height="68" rx="12" fill="{accent}"/>
  <text x="120" y="578" font-family="Arial, sans-serif" font-size="24" fill="#ffffff">Replicate this UI exactly in HTML + CSS</text>
  <text x="90" y="635" font-family="Arial, sans-serif" font-size="20" fill="#334155">{style['note']}</text>
</svg>'''

    path.write_text(svg, encoding='utf-8')

rows = []
for level in levels:
    for topic in topics:
        titles = question_templates[topic]
        for idx in range(1, 21):
            title_core = titles[idx - 1]
            title = f"{title_core} ({topic} {level} Q{idx:02d})"

            description = (
                f"Build this UI from scratch using only HTML and CSS. Produce a page that visually matches the exact reference output for {title_core}."
            )
            scenario = (
                f"An LMS challenge requires students to recreate a real UI component/page under the topic {topic}. The submission is graded on exact layout match, spacing, typography, and structure."
            )
            hint = (
                "Start with semantic structure, map each visual block from the reference, then style section-by-section: container, typography, spacing, alignment, and responsiveness."
            )
            constraint = constraints_by_level[level][(idx - 1) % len(constraints_by_level[level])]

            level_slug = slugify(level)
            topic_slug = slugify(topic)
            question_slug = f"q{idx:02d}-{slugify(title_core)}"

            svg_path = assets_root / level_slug / topic_slug / f"{question_slug}.svg"
            svg_path.parent.mkdir(parents=True, exist_ok=True)
            render_svg(svg_path, level, topic, title_core, idx)

            rel = svg_path.relative_to(repo).as_posix()
            example_output = f"https://raw.githubusercontent.com/{owner}/{repo_name}/{branch}/{rel}"

            rows.append([
                language,
                level,
                topic,
                title,
                description,
                scenario,
                hint,
                constraint,
                example_output,
            ])

wb = Workbook()
ws = wb.active
ws.title = 'Frontend Practice Bank'
headers = [
    'Language', 'Level', 'Topic', 'Title', 'Description',
    'Explanation scenario', 'Instruction hint', 'Constraints', 'Example output'
]
ws.append(headers)
for r in rows:
    ws.append(r)

for c in ['A','B','C','D','E','F','G','H','I']:
    ws.column_dimensions[c].width = 42

output = repo / 'Frontend_Development_Practice_Bank.xlsx'
wb.save(output)

print(f"Generated rows: {len(rows)}")
print(f"Saved: {output}")
