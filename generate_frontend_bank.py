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
    'HTML', 'Semantic HTML', 'Forms', 'Tables', 'CSS',
    'Box Model', 'Flexbox', 'Grid', 'Media Queries', 'Responsive Design'
]

# 30 unique prompts per topic => 300 per level
question_bank = {
    'HTML': [
        'Portfolio Hero Blueprint','Campus Club Home Structure','Startup About Storyline','Travel Blog Intro Layout','Food Menu Showcase Page',
        'Course Finder Landing Outline','NGO Mission Page Structure','Podcast Episode Landing','Fitness Coach Intro Page','Bookstore Discover Page',
        'Event Countdown Content Page','Freelancer Service Page Skeleton','Museum Exhibit Intro Page','Career Fair Info Page','Tech Meetup Announcement Page',
        'Local News Digest Home','Recipe Discovery Landing','Product Launch Intro Page','Student Project Gallery Page','Volunteer Drive Landing',
        'Language Learning Home','Internship Portal Intro','Healthcare Awareness Page','Photography Journey Page','Coding Bootcamp Info Page',
        'Town Hall Notice Page','Scholarship Program Intro','Music Festival Welcome Page','Real Estate Listing Intro','Climate Action Campaign Page'
    ],
    'Semantic HTML': [
        'Accessible Editorial Article','Landmark-rich Government Notice','Research Summary with Sections','Open Source Release Notes','Health Advisory Bulletin',
        'Newsroom Feature with Asides','Education Policy Explainer','Product Case Study Narrative','Conference Report with Time Data','Community Program Digest',
        'Public Service Announcement','Digital Magazine Story Flow','Cinema Review with Metadata','Travel Guide with Highlights','Charity Report with Figures',
        'Startup Milestone Chronicle','Interview Transcript Publication','Workshop Recap Article','Museum Research Abstract','Environment Briefing Story',
        'Tech Standards Overview','Library Initiative Report','Civic Update Thread','Agriculture Innovation Report','Sports Analysis Editorial',
        'Cultural Festival Article','Medical Procedure Overview','Urban Planning Update','Learning Roadmap Narrative','Cybersecurity Incident Summary'
    ],
    'Forms': [
        'University Admission Application','Job Referral Submission','Hospital Appointment Intake','Vacation Leave Request','Freelance Project Brief',
        'Startup Demo Registration','Scholarship Candidate Form','Event Speaker Application','Customer Complaint Intake','Subscription Upgrade Request',
        'Insurance Claim Entry','Conference Volunteer Signup','Coaching Session Booking','Product Return Request','Partnership Proposal Form',
        'Technical Support Escalation','Hackathon Team Registration','Course Feedback Survey','Housing Rental Inquiry','Internship Daily Report',
        'Bank KYC Update Form','Restaurant Catering Inquiry','NGO Volunteer Onboarding','Workshop Attendance Confirmation','SaaS Trial Activation',
        'School Parent Meeting Slot','Clinic Lab Test Booking','Mentor-Mentee Match Form','Travel Reimbursement Claim','Research Participation Consent'
    ],
    'Tables': [
        'Semester Grade Ledger','Warehouse Stock Matrix','Quarterly Sales Tracker','Hospital Shift Allocation','Flight Departure Board',
        'Retail Price Comparison','Project Sprint Burndown Table','Support SLA Performance','Attendance Compliance Sheet','Subscription Renewal Grid',
        'Website Conversion Report','Manufacturing Defect Log','Supplier Delivery Tracker','Budget vs Spend Summary','Exam Seating Allocation',
        'Sports Tournament Ladder','Library Book Circulation','Class Timetable Board','Payroll Summary Table','Incident Response Timeline',
        'Utility Usage Dashboard Table','Training Progress Matrix','Marketing Campaign Metrics','Ecommerce Order Audit','Research Experiment Log',
        'Branch-wise Revenue Table','Farm Yield Comparison','Transport Route Timings','Helpdesk Resolution Table','Customer NPS Breakdown'
    ],
    'CSS': [
        'Hero Section Theme Build','Feature Strip Visual Design','Announcement Banner Styling','Profile Card Identity System','Navigation Skin Architecture',
        'Pricing Tile Skinning','Dashboard Widget Aesthetics','Testimonial Emphasis Style','Alert Severity Theme Pack','Tag System Color Language',
        'Timeline Accent Styling','Footer Visual Hierarchy','Article Typography Refinement','Search Panel Interface Look','Promo Block Contrast Styling',
        'CTA Group Interaction Styling','Section Divider Language','Stats Counter Appearance','Form State Theme Rules','Sidebar Branding Styles',
        'Loading Block Visual State','Notification Stack Styling','Badge Priority Color Rules','Panel Depth Treatment','Card Hover Atmosphere',
        'Tab Navigation Styling','Review Block Quote Styling','Utility Class Theme Set','Feature Comparison Emphasis','Microcopy Tone Styling'
    ],
    'Box Model': [
        'Card Padding Calibration','Panel Margin Rhythm','Info Tile Border Discipline','Content Column Width Control','Widget Spacing Consistency',
        'Form Field Internal Spacing','Header-to-body Gap Control','Sidebar Item Box Sizing','Modal Frame Boundary Setup','CTA Block Breathing Space',
        'List Row Compact Spacing','Avatar Meta Box Framing','Stats Tile Edge Balance','Caption Wrapper Width Rules','Toolbar Item Separation',
        'Section Shell Spacing Scale','Promo Ribbon Box Balance','Article Paragraph Width Guard','Table Cell Padding Policy','Gallery Tile Box Balance',
        'Notification Container Sizing','Pricing Card Edge Rhythm','FAQ Panel Spacing Logic','Timeline Node Box Spacing','Action Bar Padding Grid',
        'Profile Summary Frame','Comment Box Vertical Rhythm','Checkout Summary Boundaries','Card Stack Consistent Offsets','Footer Link Block Sizing'
    ],
    'Flexbox': [
        'Top Nav Distribution','Action Toolbar Alignment','Profile Header Justification','Card Deck Wrapping Rules','Pricing Row Equalization',
        'Form Controls Inline Layout','Comment Header Alignment','Sidebar Menu Distribution','Mobile Header Flex Shift','Feature Split Justify Pattern',
        'Rating Row Alignment','Media Object Horizontal Flow','Tag Wrap Control','Checkout Totals Alignment','Icon + Label Pairing',
        'Gallery Caption Balance','Task Item Alignment','Search Result Action Row','Sticky Footer Link Alignment','KPI Row Distribution',
        'Avatar + Text Baseline','Button Cluster Wrap Strategy','Menu Toggle Spacing','Notification Item Alignment','Filter Bar Justify Rules',
        'Timeline Row Distribution','Progress Card Content Flow','List Item Meta Alignment','Chat Bubble Row Layout','Hero CTA Group Alignment'
    ],
    'Grid': [
        'Analytics Overview Lattice','Course Catalog Grid System','Portfolio Mosaic Grid','Product Shelf Grid','Editorial Magazine Grid',
        'Admin Console Grid','Learning Module Matrix','KPI Snapshot Grid','Team Directory Mosaic','News Digest Grid Frame',
        'Case Study Showcase Grid','Calendar Planning Grid','Pricing Matrix Grid','Service Card Grid','Feature Matrix Alignment Grid',
        'Sidebar + Content Grid','Blog Highlight Grid','Stats and Table Grid','Widget Control Board Grid','Research Summary Grid',
        'Image Gallery Proportion Grid','Landing Feature Grid','Operations Control Grid','Department Overview Grid','Schedule Planner Grid',
        'Offer Comparison Grid','Marketplace Listing Grid','Campaign Results Grid','Onboarding Steps Grid','Multi-panel Knowledge Grid'
    ],
    'Media Queries': [
        'Breakpoint-driven Nav Shift','Tablet Card Reflow','Desktop Typography Expansion','Form Stack Conversion','Sidebar Collapse Trigger',
        'Hero Proportion Change','Button Group Wrapping','Table-to-card Adaptation','Gallery Column Reduction','Footer Stack Breakpoint',
        'Sticky Header Enable Rule','Section Spacing Adaptation','Banner Height Tuning','CTA Position Switch','Grid Area Remap',
        'Image Treatment by Viewport','Input Width Adaptation','Pricing Row Reorder','Visibility Toggle by Width','Dashboard Block Shift',
        'Two-pane to Single-pane','List Density Adaptation','Metric Card Rearrangement','Timeline Simplification Rule','Control Panel Collapse',
        'Search Result Layout Flip','Feature Tiles Rebalance','Promo Copy Scaling','Navigation Hit Area Adjustment','Mobile-first Layering Conversion'
    ],
    'Responsive Design': [
        'SaaS Landing Multi-device Build','Education Portal Adaptive Layout','Clinic Website Fluid Flow','Restaurant Site Adaptive Home','Travel Guide Responsive Build',
        'Portfolio Site Fluid Components','Community Platform Adaptive Shell','Event Website Responsive Frame','Startup Marketing Adaptive Page','News Hub Multi-breakpoint Design',
        'Dashboard Adaptive Overview','Booking Interface Responsive Build','Magazine Site Responsive Narrative','NGO Campaign Responsive Page','Ecommerce Adaptive Catalog',
        'Help Center Responsive Structure','Developer Docs Adaptive Readability','Course Platform Responsive Cards','Music Festival Adaptive Layout','Research Portal Responsive Build',
        'Banking UI Adaptive Panel','Recruitment Portal Responsive Grid','Public Service Responsive Site','Fitness Program Adaptive Dashboard','Agriculture Insights Responsive Page',
        'Movie Review Responsive Layout','Real Estate Adaptive Listing','Logistics Tracker Responsive Panel','Tourism Board Responsive Story','Hackathon Hub Adaptive Framework'
    ]
}

level_constraints = {
    'Beginner': [
        'Use only HTML and CSS; no JavaScript.',
        'Match visible spacing, sizes, and alignment exactly.',
        'Use meaningful semantic tags for content blocks.',
        'Do not use external CSS frameworks or libraries.',
        'Keep class names clear and beginner-readable.'
    ],
    'Intermediate': [
        'Implement mobile-first CSS with at least two breakpoints.',
        'Add keyboard-visible focus states on interactive elements.',
        'Use reusable component classes and avoid duplication.',
        'Maintain consistent spacing scale across sections.',
        'Do not rely on absolute positioning unless necessary.'
    ],
    'Hard': [
        'Support mobile, tablet, and desktop without layout breakage.',
        'Ensure accessible contrast and clear focus indicators.',
        'Use low-specificity, maintainable CSS architecture.',
        'Prevent cumulative layout shift between breakpoints.',
        'Deliver production-quality structure and naming.'
    ]
}

contexts = [
    'LMS timed frontend challenge', 'UI recreation assessment', 'placement-prep coding round', 'design-to-code exercise',
    'accessibility-first implementation drill', 'responsive layout checkpoint', 'component architecture practice',
    'semantic and structure evaluation', 'pixel-accuracy assignment', 'production-readiness practice'
]

hints = [
    'Sketch the page blocks first, then code structure before styling.',
    'Start with a parent container and map each reference segment to a semantic section.',
    'Apply typography and spacing tokens early to keep visual rhythm consistent.',
    'Build mobile layout first, then scale up using media queries.',
    'Use reusable utility/component classes instead of repeating declarations.',
    'Validate alignment with consistent gap/padding increments.',
    'Keep source order logical before handling visual arrangements.',
    'Test at multiple widths after each major section implementation.',
    'Match card/section proportions before fine-tuning colors and shadows.',
    'Add focus styles and verify keyboard navigation behavior.'
]

level_visual = {
    'Beginner': ('#0f766e', '#f0fdfa'),
    'Intermediate': ('#1d4ed8', '#eff6ff'),
    'Hard': ('#b45309', '#fff7ed')
}

def slugify(text: str) -> str:
    return re.sub(r'[^a-z0-9]+', '-', text.lower()).strip('-')

def make_svg(path: Path, level: str, topic: str, title: str, idx: int):
    accent, bg = level_visual[level]
    v = idx % 5
    blocks = [
        '<rect x="88" y="210" width="1050" height="90" rx="14" fill="#dbeafe"/>\n<rect x="88" y="320" width="515" height="190" rx="14" fill="#e2e8f0"/>\n<rect x="623" y="320" width="515" height="190" rx="14" fill="#fef3c7"/>',
        '<rect x="88" y="210" width="330" height="300" rx="14" fill="#e2e8f0"/>\n<rect x="438" y="210" width="700" height="140" rx="14" fill="#dbeafe"/>\n<rect x="438" y="370" width="340" height="140" rx="14" fill="#fef3c7"/>\n<rect x="798" y="370" width="340" height="140" rx="14" fill="#ede9fe"/>',
        '<rect x="88" y="210" width="250" height="140" rx="14" fill="#dbeafe"/>\n<rect x="358" y="210" width="250" height="140" rx="14" fill="#e2e8f0"/>\n<rect x="628" y="210" width="250" height="140" rx="14" fill="#fef3c7"/>\n<rect x="898" y="210" width="240" height="140" rx="14" fill="#ede9fe"/>\n<rect x="88" y="370" width="1050" height="140" rx="14" fill="#f1f5f9"/>',
        '<rect x="88" y="210" width="1050" height="140" rx="14" fill="#f1f5f9"/>\n<rect x="88" y="370" width="340" height="140" rx="14" fill="#dbeafe"/>\n<rect x="448" y="370" width="340" height="140" rx="14" fill="#e2e8f0"/>\n<rect x="808" y="370" width="330" height="140" rx="14" fill="#fef3c7"/>',
        '<rect x="88" y="210" width="510" height="300" rx="14" fill="#e2e8f0"/>\n<rect x="618" y="210" width="520" height="90" rx="14" fill="#dbeafe"/>\n<rect x="618" y="320" width="520" height="90" rx="14" fill="#fef3c7"/>\n<rect x="618" y="430" width="520" height="80" rx="14" fill="#ede9fe"/>'
    ][v]
    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <rect width="1280" height="720" fill="{bg}"/>
  <rect x="48" y="48" width="1184" height="624" rx="24" fill="#ffffff" stroke="{accent}" stroke-width="5"/>
  <text x="88" y="108" font-family="Arial, sans-serif" font-size="34" font-weight="700" fill="{accent}">{title}</text>
  <text x="88" y="152" font-family="Arial, sans-serif" font-size="22" fill="#0f172a">{level} • {topic} • Reference UI #{idx:03d}</text>
  {blocks}
  <rect x="88" y="535" width="520" height="66" rx="12" fill="{accent}"/>
  <text x="116" y="577" font-family="Arial, sans-serif" font-size="23" fill="#ffffff">Code this UI exactly in HTML + CSS</text>
</svg>'''
    path.write_text(svg, encoding='utf-8')

columns = [
    'Language','Level','Topic','Title','Description','Explanation scenario','Instruction hint','Constraints','Example output'
]

wb = Workbook()
wb.remove(wb.active)

for level in levels:
    ws = wb.create_sheet(level)
    ws.append(columns)
    row_idx = 0
    for topic in topics:
        prompts = question_bank[topic]
        for local_idx, prompt in enumerate(prompts, start=1):
            row_idx += 1
            global_id = row_idx
            title = f"{prompt} - {topic} {level} Challenge {global_id:03d}"
            description = (
                f"Build from scratch a {topic.lower()} problem titled '{prompt}'. Create the complete UI so that it matches the reference output exactly in structure, spacing, hierarchy, and styling."
            )
            scenario = (
                f"This problem is framed as a {contexts[(global_id-1) % len(contexts)]}. Students are evaluated on how accurately they transform requirements into a working frontend implementation for the '{prompt}' use case."
            )
            instruction = (
                f"Implement the '{prompt}' interface using semantic HTML and CSS. Recreate every major block from the reference image, then refine typography, spacing, alignment, and responsive behavior to achieve a pixel-close match."
            )
            constraint = (
                f"{level_constraints[level][(global_id-1) % len(level_constraints[level])]} Additional rule: ensure this solution is unique to challenge {global_id:03d} and avoids copying layout code from earlier tasks."
            )

            lvl_slug = slugify(level)
            topic_slug = slugify(topic)
            q_slug = slugify(prompt)
            svg_path = assets_root / lvl_slug / topic_slug / f"q{global_id:03d}-{q_slug}.svg"
            svg_path.parent.mkdir(parents=True, exist_ok=True)
            make_svg(svg_path, level, topic, prompt, global_id)

            rel = svg_path.relative_to(repo).as_posix()
            link = f"https://raw.githubusercontent.com/{owner}/{repo_name}/{branch}/{rel}"

            ws.append([
                language, level, topic, title, description,
                scenario, instruction + ' Hint: ' + hints[(global_id-1) % len(hints)], constraint, link
            ])

    for col in ['A','B','C','D','E','F','G','H','I']:
        ws.column_dimensions[col].width = 42

out = repo / 'Frontend_Development_Practice_Bank.xlsx'
wb.save(out)
print('Generated workbook with 3 sheets x 300 questions each (900 total).')
print(out)
