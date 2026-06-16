from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
from pathlib import Path

OUT = Path(r"C:\Users\Kyle\.local\bin\logistics_auditor\Logistics_Margin_Diagnostic_Marketing_Strategy.docx")

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def x(s):
    return escape(str(s), {'"': '&quot;'})


def run(text, bold=False, italic=False, size=22, color=None):
    props = []
    if bold:
        props.append('<w:b/>')
    if italic:
        props.append('<w:i/>')
    if size:
        props.append(f'<w:sz w:val="{size}"/>')
        props.append(f'<w:szCs w:val="{size}"/>')
    if color:
        props.append(f'<w:color w:val="{color}"/>')
    rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    return f"<w:r>{rpr}<w:t xml:space=\"preserve\">{x(text)}</w:t></w:r>"


def para(text="", style=None, align=None, bold=False, italic=False, size=22, color=None, spacing_after=120):
    ppr = []
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    ppr.append(f'<w:spacing w:after="{spacing_after}"/>')
    return f"<w:p><w:pPr>{''.join(ppr)}</w:pPr>{run(text, bold, italic, size, color)}</w:p>"


def bullets(items):
    return ''.join(para(f"• {item}", spacing_after=80) for item in items)


def numbered(items):
    return ''.join(para(f"{i}. {item}", spacing_after=80) for i, item in enumerate(items, 1))


def heading(text, level=1):
    style = f"Heading{level}"
    return para(text, style=style, bold=True, size={1:32,2:28,3:24}.get(level,22), color="1F4E79", spacing_after=160)


def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def table(rows, widths=None):
    if widths is None:
        widths = [3000] * len(rows[0])
    xml = ['<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>']
    for width in widths:
        xml.append(f'<w:gridCol w:w="{width}"/>')
    xml.append('</w:tblGrid>')
    for r_idx, row in enumerate(rows):
        xml.append('<w:tr>')
        for cell in row:
            shade = '<w:shd w:fill="D9EAF7"/>' if r_idx == 0 else ''
            xml.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths[0]}" w:type="dxa"/>{shade}</w:tcPr>')
            for line in str(cell).split('\n'):
                xml.append(para(line, bold=(r_idx == 0), size=20, spacing_after=60))
            xml.append('</w:tc>')
        xml.append('</w:tr>')
    xml.append('</w:tbl>')
    return ''.join(xml) + para("", spacing_after=120)


parts = []
parts.append(para("Strategic Marketing Strategy", align="center", bold=True, size=40, color="1F4E79", spacing_after=80))
parts.append(para("for the 360° Margin Diagnostic Engine", align="center", bold=True, size=30, color="2F5597", spacing_after=220))
parts.append(para("A founder-led, tech-enabled logistics margin recovery and governance venture for South African e-commerce brands", align="center", italic=True, size=22, spacing_after=420))
parts.append(para("Prepared for: Venture validation, founder outreach, and university capstone presentation", align="center", size=21, spacing_after=100))
parts.append(para("Founder context: 21-year-old final-year BCom Management Sciences student at Stellenbosch University, specialising in Business Analytics, Logistics, and Supply Chain Management", align="center", size=21, spacing_after=100))
parts.append(para("Business model: Free diagnostic audits during validation, converting into monthly Logistics Spend Governance retainers", align="center", size=21, spacing_after=500))
parts.append(para("Confidential strategic working document", align="center", italic=True, size=18, color="666666"))
parts.append(page_break())

parts.append(heading("1. Executive Summary", 1))
parts.append(para("The 360° Margin Diagnostic Engine is a specialised analytics and advisory offering for mid-market South African e-commerce brands. It uses a Python-based Streamlit application to ingest chaotic courier CSV exports, clean the data, identify logistics margin leakage, and generate a professional Recovery Blueprint that founders can use to recover overcharges and improve warehouse execution."))
parts.append(para("The commercial opportunity exists because growing e-commerce brands often know that courier costs are painful, but they do not know exactly which waybills, packaging decisions, SKUs, or shipping zones are causing the leakage. Couriers and shipping platforms provide shipment execution, but they are not economically incentivised to act as independent margin auditors. Internal finance teams are usually too busy to audit thousands of line items manually."))
parts.append(para("This venture should therefore position itself as an independent logistics margin auditor: a practical, data-led partner that turns messy operational exports into recoverable margin, stronger warehouse SOPs, and continuous cost governance."))
parts.append(table([
    ["Strategic question", "Recommended answer"],
    ["What is being sold?", "Not merely a dashboard, but a decision-ready logistics margin audit and Recovery Blueprint."],
    ["Who is it for?", "Mid-market South African e-commerce brands with meaningful courier spend and limited internal audit capacity."],
    ["Why now?", "Inflation, fuel volatility, competitive e-commerce pricing, and rising fulfilment costs make hidden margin leakage more damaging."],
    ["Why you?", "You combine Stellenbosch logistics and analytics training with hands-on Shopify and B2B sales experience."],
    ["How should it make money?", "Start with free audits to prove value; convert successful audits into monthly Logistics Spend Governance retainers."]
], [2800, 6200]))

parts.append(heading("2. Product and Business Synthesis", 1))
parts.append(heading("2.1 Product Definition", 2))
parts.append(para("The product is a 360° Margin Diagnostic Engine that audits e-commerce logistics spend across three leakage pillars:"))
parts.append(bullets([
    "Courier billing anomalies: deterministic checks that identify waybills where billed volumetric weight appears materially inconsistent with physical dimensions or expected shipment logic.",
    "Packaging inefficiencies: detection of the Empty Box Tax, where oversized packaging inflates volumetric billing, material cost, and warehouse complexity.",
    "SKU velocity and trapped capital: cross-analysis of SKU movement, basket patterns, and zone-level shipping cost to identify products that appear profitable before fulfilment but destroy margin after delivery."
]))
parts.append(heading("2.2 Business Model", 2))
parts.append(para("The business model is a tech-enabled agency that uses proprietary analytics to deliver a high-trust advisory output. The initial 30-day validation phase should offer free audits to collect proof, generate testimonials, and validate the strength of the leakage hypothesis. The long-term model should shift to a monthly retainer called Logistics Spend Governance."))
parts.append(heading("2.3 Core Strategic Concept: Arbitrage of Friction", 2))
parts.append(para("The business creates value through an arbitrage of friction. Courier data is messy, billing rules are technical, warehouse packaging behaviour is inconsistent, and founders have limited time to investigate every shipment. The venture captures value by doing what the merchant knows is important but cannot realistically do manually: convert fragmented logistics data into specific, quantified, action-ready margin improvements."))

parts.append(heading("3. Brand Positioning Strategy", 1))
parts.append(heading("3.1 Recommended Positioning Statement", 2))
parts.append(para("We help South African e-commerce brands recover hidden profit lost through courier billing anomalies, oversized packaging, and SKU-level logistics inefficiencies by turning messy courier data into a practical Recovery Blueprint and monthly spend governance system.", italic=True, size=24, color="1F4E79"))
parts.append(heading("3.2 Founder Positioning as a 21-Year-Old", 2))
parts.append(para("Your age should not be hidden, but it should not be the centre of the pitch. The strongest positioning is not “young founder asking for a chance”; it is “specialised student-operator with uncommon focus, technical capability, and fresh analytical energy.” In a market where many agencies sell generic growth services, your advantage is that you are narrowly focused on a painful operational problem and have already built the diagnostic engine."))
parts.append(table([
    ["Potential concern", "How to neutralise it"],
    ["He is young", "Lead with the audit output, sample report, and evidence of savings rather than biography."],
    ["Is this professional enough?", "Use a clean PDF, confidentiality wording, structured data request, and clear assumptions."],
    ["Can he understand my business?", "Reference your Shopify operating experience and focus on e-commerce-specific logistics pain."],
    ["Can I trust the data handling?", "Offer NDAs, request only necessary fields, anonymise benchmarks, and explain data deletion procedures."],
    ["Will this take too much effort?", "Make the process simple: export CSV, send file, receive Recovery Blueprint."]
], [2600, 6400]))
parts.append(heading("3.3 Brand Personality", 2))
parts.append(bullets([
    "Analytical, not theoretical.",
    "Founder-friendly, not corporate and inaccessible.",
    "Independent from couriers and therefore commercially aligned with the merchant.",
    "Precise and evidence-led, avoiding exaggerated savings claims.",
    "Local to South Africa and aware of the country’s logistics realities."
]))

parts.append(heading("4. Ideal Customer Profile and Segmentation", 1))
parts.append(para("The best early customers are not micro-stores with low volume and not large enterprises with formal supply chain departments. The strongest wedge is the mid-market brand: large enough for logistics leakage to be material, but not yet mature enough to have a dedicated logistics audit function."))
parts.append(table([
    ["Segment", "Fit", "Reason"],
    ["Small Shopify stores under 100 shipments/month", "Low", "Leakage may exist, but savings are often too small to justify advisory time."],
    ["Mid-market brands with 500–10,000 shipments/month", "High", "Courier spend is meaningful, data exists, and founders still care directly about margin."],
    ["Enterprise retailers", "Medium", "Large opportunity, but slower sales cycles and stronger internal teams."],
    ["Agencies and fulfilment partners", "Medium-high", "Could become channel partners if the offer helps their clients reduce cost."],
    ["Couriers", "Low as clients", "Incentives conflict with independent overbilling and margin audit positioning."]
], [2600, 1300, 5100]))
parts.append(heading("4.1 Best Early Vertical Targets", 2))
parts.append(bullets([
    "Fashion and apparel brands: high order frequency, returns exposure, packaging variation, and competitive delivery expectations.",
    "Beauty, skincare, and cosmetics brands: small items often shipped inefficiently in oversized boxes.",
    "Supplements and health products: repeat purchases make monthly logistics governance valuable.",
    "Pet products and homeware: dimensional weight and bulky SKUs can create significant margin leakage.",
    "Premium DTC brands: more likely to value professional reporting and margin protection."
]))

parts.append(heading("5. Value Proposition and Messaging", 1))
parts.append(heading("5.1 Primary Value Proposition", 2))
parts.append(para("Your courier bill is not just an expense; it is an unaudited margin system. We identify where that system is leaking and give you a Recovery Blueprint to recover overcharges, reduce packaging waste, and protect contribution margin.", italic=True, size=24, color="1F4E79"))
parts.append(heading("5.2 Pain-Based Messaging Angles", 2))
parts.append(table([
    ["Pain", "Message"],
    ["Courier bill keeps rising", "Find the shipments, zones, and packaging decisions driving the increase."],
    ["Founder lacks time", "Send the exports; receive a decision-ready Recovery Blueprint."],
    ["Finance only checks invoice totals", "We audit the waybill-level logic behind the total."],
    ["Warehouse uses oversized boxes", "Quantify the Empty Box Tax and create packaging rules that protect margin."],
    ["Some SKUs sell but do not make money", "Identify dead-stock SKUs that bleed margin after shipping cost."],
    ["Couriers control the data", "Use independent evidence to dispute charges and negotiate from a stronger position."]
], [3100, 5900]))
parts.append(heading("5.3 Messaging Principles", 2))
parts.append(bullets([
    "Do not sell software first. Sell the financial outcome: recovered margin and operational control.",
    "Avoid claiming guaranteed savings. Say: “We identify and quantify likely leakage.”",
    "Use specific local references: Bob Go, The Courier Guy, courier CSVs, volumetric billing, zone costs, free shipping thresholds.",
    "Frame the audit as low-effort for the founder.",
    "Make the report feel boardroom-ready even if the analysis begins with messy CSV files."
]))

parts.append(heading("6. Offer Architecture", 1))
parts.append(table([
    ["Offer", "Purpose", "Deliverables", "Commercial role"],
    ["Free Diagnostic Audit", "Validate pain and acquire proof", "Data import, anomaly scan, Recovery Blueprint summary, savings estimate", "Lead magnet and trust builder"],
    ["Recovery Blueprint", "Convert analysis into action", "PDF report, dispute list, packaging insights, SKU/zone leakage analysis", "Main proof artifact"],
    ["Implementation Sprint", "Help client apply findings", "Packaging matrix, courier dispute pack, SKU recommendations", "Paid bridge offer"],
    ["Monthly Logistics Spend Governance", "Prevent leakage from returning", "Monthly audit, anomaly report, packaging compliance, SKU margin review", "Recurring revenue model"],
    ["Rate Negotiation Support", "Increase strategic value", "Courier benchmarking, rate-card analysis, negotiation pack", "Premium advisory upsell"]
], [2000, 2100, 3300, 1600]))
parts.append(heading("6.1 Suggested Retainer Packages", 2))
parts.append(table([
    ["Package", "Indicative positioning", "Scope"],
    ["Starter Governance", "For growing stores wanting monthly invoice control", "Monthly courier anomaly audit, top leakage report, dispute-ready export."],
    ["Growth Governance", "For brands with meaningful order volume and packaging complexity", "Everything in Starter plus packaging compliance, SKU-level logistics margin review, and founder summary."],
    ["Strategic Governance", "For larger brands preparing to renegotiate courier rates", "Everything in Growth plus quarterly rate benchmarking, courier performance analysis, and negotiation pack."]
], [2200, 2700, 4100]))
parts.append(para("Pricing should only be finalised after validation data shows average recoverable value. As a principle, the monthly retainer should feel small relative to identified monthly leakage and the cost of ignoring it."))

parts.append(heading("7. Go-to-Market Strategy", 1))
parts.append(heading("7.1 30-Day Validation Objective", 2))
parts.append(para("The goal of the first 30 days is not to maximise revenue. The goal is to prove that the problem is common, measurable, and urgent enough for founders to act on."))
parts.append(table([
    ["Validation metric", "Target"],
    ["Qualified brands contacted", "150–300"],
    ["Founder/operator replies", "15–30"],
    ["Completed audits", "5–10"],
    ["Case studies/testimonials", "2–3"],
    ["Average leakage identified", "Track as rand value and percentage of courier spend"],
    ["Retainer conversion conversations", "At least 3 serious discussions"]
], [3600, 5400]))
parts.append(heading("7.2 Outreach Channels", 2))
parts.append(bullets([
    "Cold email to founders, operations managers, and finance leads of South African e-commerce brands.",
    "LinkedIn founder outreach using a concise, problem-led message.",
    "Partnership outreach to Shopify agencies, fulfilment providers, and e-commerce accountants.",
    "Founder communities and South African e-commerce groups, using educational posts rather than aggressive selling.",
    "Direct Loom-style audit teaser videos for high-value prospects once a public shipping or product pattern suggests likely leakage."
]))
parts.append(heading("7.3 Cold Email Framework", 2))
parts.append(para("Subject options:"))
parts.append(bullets([
    "Possible courier margin leakage in your fulfilment data",
    "Free logistics margin audit for [Brand]",
    "Are oversized boxes inflating your courier bill?",
    "I built a tool to audit SA e-commerce courier spend"
]))
parts.append(para("Email template:"))
parts.append(para("Hi [Name],\n\nI’m a final-year Management Sciences student at Stellenbosch focusing on logistics, supply chain, and business analytics. I also run e-commerce projects myself, so I know how easily courier costs can quietly eat margin.\n\nI built a Python-based diagnostic engine that audits messy courier CSV exports from platforms like Bob Go and The Courier Guy. It flags three things: possible volumetric billing anomalies, oversized packaging decisions, and SKUs that lose margin once delivery cost is included.\n\nI’m currently offering a small number of free audits to validate the model. The output is a short Recovery Blueprint showing where margin may be leaking and what to fix. No platform access is needed — only exported courier/order CSVs.\n\nWould you be open to me running a free audit for [Brand]?\n\nBest,\n[Your Name]"))
parts.append(heading("7.4 Sales Process", 2))
parts.append(numbered([
    "Identify target brands with visible e-commerce activity and likely courier spend.",
    "Send a short, specific outreach message focused on courier leakage, not software.",
    "Offer the free audit as a low-risk diagnostic.",
    "Collect CSV exports using a clear data request checklist.",
    "Deliver the Recovery Blueprint in a professional PDF format.",
    "Walk the founder through the top three findings in a 20-minute call.",
    "Convert the conversation to either an implementation sprint or monthly governance retainer."
]))

parts.append(heading("8. Marketing Funnel", 1))
parts.append(table([
    ["Funnel stage", "Objective", "Content or action"],
    ["Awareness", "Make founders aware courier leakage exists", "LinkedIn posts: volumetric billing, Empty Box Tax, free shipping margin trap."],
    ["Interest", "Show that the problem is measurable", "Sample screenshots, anonymised audit snippets, mini case studies."],
    ["Consideration", "Build trust and reduce perceived effort", "Free audit offer, NDA option, data checklist, sample Recovery Blueprint."],
    ["Conversion", "Get data and deliver proof", "Audit delivery, founder review call, quantified savings."],
    ["Retention", "Shift from project to governance", "Monthly report, recurring anomaly tracking, packaging compliance dashboard."],
    ["Referral", "Turn clients into proof engines", "Testimonials, anonymised savings, founder introductions."]
], [2100, 2600, 4300]))
parts.append(heading("8.1 Content Themes", 2))
parts.append(bullets([
    "The Empty Box Tax: how oversized packaging quietly inflates courier bills.",
    "Why checking invoice totals is not the same as auditing courier spend.",
    "The free shipping trap: when delivery promises destroy contribution margin.",
    "Volumetric weight explained for e-commerce founders.",
    "Dead stock is worse when it is expensive to ship.",
    "Why mid-market brands need logistics spend governance before they scale further."
]))

parts.append(heading("9. SWOT Analysis", 1))
parts.append(table([
    ["Strengths", "Weaknesses"],
    ["Working technical pipeline; strong niche focus; low overhead; academic credibility; Shopify and B2B sales experience; quantifiable ROI.", "Reliance on messy client data; limited initial trust; manual cleaning burden; savings attribution may be challenged; early brand has limited proof."],
    ["Opportunities", "Threats"],
    ["Rate negotiation; monthly governance; anonymised benchmark dataset; warehouse SOP consulting; Shopify/courier API integrations; vertical-specific playbooks.", "Couriers improving tooling; clients fixing problems internally; SaaS platforms adding analytics; data access resistance; low urgency during growth phases."]
], [4500, 4500]))

parts.append(heading("10. Market Environment Analysis", 1))
parts.append(para("The immediate competitive environment consists mainly of indirect alternatives. SaaS analytics platforms often focus on sales, inventory, and customer metrics but rarely audit waybill-level courier logic. Couriers and shipping aggregators provide shipment execution and basic reporting, but their incentives do not fully align with reducing the merchant’s bill. In-house finance teams may reconcile totals but usually lack time and operational context to investigate individual shipments. Traditional consultants may offer supply chain advice, but they are often too expensive or too enterprise-focused for mid-market e-commerce brands."))
parts.append(para("The key power dynamic is information asymmetry. Couriers have scale, standard billing rules, operational data, and control over scans. Mid-market merchants have limited leverage and fragmented information. This business helps rebalance that asymmetry by giving merchants evidence, quantified leakage, and a structured basis for disputes, SOP changes, and negotiations."))

parts.append(heading("11. PESTLE Analysis: South Africa", 1))
parts.append(table([
    ["Factor", "Implication for the business"],
    ["Political", "Infrastructure pressure and uneven service delivery increase the need for private logistics resilience and cost visibility."],
    ["Economic", "Inflation, fuel volatility, currency pressure, and consumer price sensitivity make margin recovery commercially urgent."],
    ["Social", "Consumers expect affordable, reliable delivery, forcing merchants to absorb or optimise fulfilment costs."],
    ["Technological", "Growth in Shopify, WooCommerce, shipping platforms, APIs, and CSV exports makes data-led auditing increasingly feasible."],
    ["Legal", "POPIA and commercial confidentiality require disciplined data handling, NDAs, minimisation, and anonymisation."],
    ["Environmental", "Oversized packaging increases waste and transport inefficiency; reducing the Empty Box Tax supports both profit and sustainability."]
], [2000, 7000]))

parts.append(heading("12. 90-Day Execution Roadmap", 1))
parts.append(table([
    ["Period", "Focus", "Actions"],
    ["Days 1–30", "Validation", "Run free audits, test outreach, collect leakage evidence, refine data templates, build sample Recovery Blueprint."],
    ["Days 31–60", "Credibility", "Publish anonymised findings, collect testimonials, formalise offer packages, improve report design, create onboarding checklist."],
    ["Days 61–90", "Monetisation", "Convert best-fit audit clients into retainers, test implementation sprint pricing, approach agencies and fulfilment partners."],
    ["Beyond 90 days", "Scale", "Build courier-specific adapters, benchmark database, API roadmap, and repeatable monthly governance workflow."]
], [1800, 2200, 5000]))

parts.append(heading("13. Key Performance Indicators", 1))
parts.append(bullets([
    "Outreach reply rate by segment and message angle.",
    "Percentage of prospects willing to share data for a free audit.",
    "Average rand value of suspected leakage identified per audit.",
    "Leakage as a percentage of monthly courier spend.",
    "Number of dispute-ready anomalies produced per audit.",
    "Number of packaging SOP recommendations per client.",
    "Audit-to-call conversion rate.",
    "Audit-to-retainer conversion rate.",
    "Monthly recurring revenue once retainers begin.",
    "Time required to clean and process each client’s data."
]))

parts.append(heading("14. Risk Mitigation", 1))
parts.append(table([
    ["Risk", "Mitigation"],
    ["Clients hesitate to share data", "Offer NDA, request limited fields, provide sample report, explain deletion/anonymisation process."],
    ["CSV formats vary heavily", "Create a canonical schema and courier-specific import adapters."],
    ["Savings are estimates", "Separate confirmed anomalies, probable savings, and strategic opportunities in the report."],
    ["Founder age reduces perceived credibility", "Lead with professionalism, evidence, academic fit, and the quality of the Recovery Blueprint."],
    ["Clients fix once-off issues and churn", "Sell monthly governance as the control system that prevents leakage from returning."],
    ["Couriers improve their tools", "Differentiate beyond billing: packaging, SKU economics, governance, and independent advisory."
]], [3000, 6000]))

parts.append(heading("15. Final Strategic Recommendation", 1))
parts.append(para("The strongest strategy is to position the venture as a focused, independent logistics margin audit and governance partner for South African e-commerce brands. The marketing should not overemphasise the software interface; it should emphasise the founder’s problem: courier costs are rising, margins are tightening, and operational leakage is hidden inside messy data."))
parts.append(para("As a 21-year-old founder, the most credible route is not to pretend to be a large consultancy. Instead, the strategy should be to look unusually specialised, technically capable, academically grounded, and commercially practical. Your youth becomes an advantage when framed as focus, speed, and modern analytical capability — but only if the client-facing materials are polished, specific, and evidence-led."))
parts.append(para("The immediate objective is to complete enough free audits to prove that leakage is common and financially meaningful. Once proof exists, the business should shift the conversation from “I found savings once” to “I can make sure this does not happen again every month.” That is the bridge from a student-built diagnostic tool to a serious recurring logistics spend governance business."))

settings = '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/></w:settings>'
styles = '''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/></w:rPr><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="240" w:after="160"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="200" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5597"/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="160" w:after="100"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5597"/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders></w:tblPr></w:style>
</w:styles>'''

doc = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {''.join(parts)}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

with ZipFile(OUT, 'w', ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc)
    z.writestr('word/styles.xml', styles)
    z.writestr('word/settings.xml', settings)
    z.writestr('word/_rels/document.xml.rels', doc_rels)

print(str(OUT))
