from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
from pathlib import Path

OUT = Path(r"C:\Users\Kyle\.local\bin\logistics_auditor\Logistics_Auditor_Pricing_Strategy.docx")


def x(value):
    return escape(str(value), {'"': '&quot;'})


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
    return f'<w:r>{rpr}<w:t xml:space="preserve">{x(text)}</w:t></w:r>'


def para(text="", style=None, align=None, bold=False, italic=False, size=22, color=None, spacing_after=120):
    ppr = []
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    ppr.append(f'<w:spacing w:after="{spacing_after}"/>')
    return f"<w:p><w:pPr>{''.join(ppr)}</w:pPr>{run(text, bold, italic, size, color)}</w:p>"


def heading(text, level=1):
    return para(
        text,
        style=f"Heading{level}",
        bold=True,
        size={1: 32, 2: 28, 3: 24}.get(level, 22),
        color="1F4E79",
        spacing_after=160,
    )


def bullets(items):
    return ''.join(para(f"• {item}", spacing_after=80) for item in items)


def numbered(items):
    return ''.join(para(f"{i}. {item}", spacing_after=80) for i, item in enumerate(items, 1))


def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def table(rows, widths=None):
    if widths is None:
        widths = [2500] * len(rows[0])
    xml = [
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        '</w:tblPr><w:tblGrid>'
    ]
    for width in widths:
        xml.append(f'<w:gridCol w:w="{width}"/>')
    xml.append('</w:tblGrid>')
    for r_idx, row in enumerate(rows):
        xml.append('<w:tr>')
        for c_idx, cell in enumerate(row):
            shade = '<w:shd w:fill="D9EAF7"/>' if r_idx == 0 else ''
            width = widths[c_idx] if c_idx < len(widths) else widths[-1]
            xml.append(f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shade}</w:tcPr>')
            for line in str(cell).split('\n'):
                xml.append(para(line, bold=(r_idx == 0), size=19 if r_idx == 0 else 18, spacing_after=40))
            xml.append('</w:tc>')
        xml.append('</w:tr>')
    xml.append('</w:tbl>')
    return ''.join(xml) + para("", spacing_after=120)


parts = []
parts.append(para("Pricing Strategy", align="center", bold=True, size=44, color="1F4E79", spacing_after=80))
parts.append(para("for a South African E-Commerce Logistics Auditing Business", align="center", bold=True, size=28, color="2F5597", spacing_after=220))
parts.append(para("A hybrid model for historical margin recovery and durable monthly logistics governance", align="center", italic=True, size=22, spacing_after=420))
parts.append(para("Prepared for: Founder-led logistics margin recovery and diagnostic engine", align="center", size=20, spacing_after=80))
parts.append(para("Currency: South African Rand (ZAR)", align="center", size=20, spacing_after=500))

parts.append(heading("Executive Recommendation", 1))
parts.append(para("The business should move from a free validation model into a hybrid pricing architecture: a premium historical audit in Month 1, followed by a predictable monthly governance retainer from Month 2 onward."))
parts.append(para("The commercial principle is simple:", bold=True))
parts.append(para("Month 1 is a forensic recovery event. Month 2 onward is logistics margin governance.", bold=True, color="1F4E79"))
parts.append(para("This prevents the success trap where excellent operational fixes reduce future performance fees to zero. The client is not only paying for mistakes found; they are paying for a monthly control system that keeps courier costs, packaging discipline, invoice accuracy, and margin leakage under governance as order volume changes."))

parts.append(table([
    ["Component", "Recommended Model", "Purpose"],
    ["Month 1: Historical Audit", "Fixed audit fee + success fee", "Capture initial value while protecting delivery time"],
    ["Month 2+: Governance", "Monthly retainer by courier spend", "Create predictable recurring revenue"],
    ["Performance Upside", "Small success fee above threshold", "Keep upside without depending on leaks"],
    ["Guarantee", "Credit shortfall, not cash refund", "Reduce founder risk while protecting cash flow"],
], [2200, 3000, 4200]))

parts.append(page_break())
parts.append(heading("1. The Bounty Hunter Tier: Month 1 Historical Audit", 1))
parts.append(heading("Recommended Structure", 2))
parts.append(para("Use a hybrid model: a fixed audit fee plus a success fee on verified recoveries or first-60-day packaging savings."))
parts.append(para("A pure percentage-of-recovery model should be avoided because it undervalues the diagnostic IP, transfers too much delivery risk to the agency, and causes revenue to decline as the client becomes operationally cleaner."))

parts.append(heading("Historical Audit Pricing", 2))
parts.append(table([
    ["Client Monthly Courier Spend", "Once-Off Audit Fee", "Success Fee", "Recommended Use"],
    ["Under R50,000", "R9,500", "25%", "Small growing brands / early case studies"],
    ["R50,000–R150,000", "R18,000", "25%", "Entry-level commercial offer"],
    ["R150,000–R500,000", "R32,000", "22.5%", "Default flagship mid-market offer"],
    ["R500,000+", "R55,000", "20%", "Larger brands with operational complexity"],
    ["R1,000,000+", "R85,000", "15%–18%", "Enterprise-level custom quote"],
], [2500, 2000, 1600, 3400]))

parts.append(heading("What Counts as Verified Value", 2))
parts.append(bullets([
    "Courier credits issued or refunds recovered.",
    "Duplicate charges reversed.",
    "Incorrect weight, volumetric, zone, fuel, surcharge, or service-level charges corrected.",
    "Packaging changes with measurable courier cost reduction over the first 60 days.",
    "Documented recoveries or avoided charges that can be tied to the diagnostic work.",
]))
parts.append(para("Avoid charging a success fee on vague or theoretical opportunities. Charge only on concrete, defensible financial impact."))

parts.append(heading("Default Month 1 Sales Line", 2))
parts.append(para("The historical audit is R32,000 once-off, plus 22.5% of verified courier credits, refunds, and first-60-day packaging savings. If we find nothing meaningful, there is no success fee.", bold=True, color="1F4E79"))

parts.append(page_break())
parts.append(heading("2. The Governance Tier: Month 2+ Retainer", 1))
parts.append(para("The retainer should be priced by monthly courier spend, not by how many errors are found. The client is buying logistics cost control, invoice governance, packaging discipline, rate-card enforcement, and founder-level visibility."))

parts.append(heading("Tiered Monthly Subscription Pricing", 2))
parts.append(table([
    ["Tier", "Monthly Courier Spend", "Monthly Retainer", "Optional Upside Fee"],
    ["Tier 1: Courier Control", "Under R50,000", "R7,500/month", "10% of verified recoveries above R15,000/month"],
    ["Tier 2: Margin Protection", "R50,000–R150,000", "R14,500/month", "10% of verified recoveries above R30,000/month"],
    ["Tier 3: Logistics Governance", "R150,000+", "R29,000/month", "7.5% of verified recoveries above R60,000/month"],
], [2100, 2500, 2300, 3400]))

parts.append(heading("Tier 1: Courier Control — R7,500/month", 2))
parts.append(para("Best for brands spending under R50,000 per month on courier costs."))
parts.append(bullets([
    "Monthly courier export review.",
    "Basic overcharge and anomaly detection.",
    "Weight and volumetric charge checks.",
    "Surcharge review.",
    "Basic packaging inefficiency review.",
    "Monthly leakage report.",
    "Dispute pack preparation.",
    "One monthly email summary or short call.",
]))

parts.append(heading("Tier 2: Margin Protection — R14,500/month", 2))
parts.append(para("Best for brands spending R50,000–R150,000 per month on courier costs. This is likely to be the bread-and-butter tier for founder-led brands."))
parts.append(bullets([
    "Everything in Tier 1.",
    "Monthly invoice audit.",
    "Weekly anomaly review.",
    "Packaging recommendation report.",
    "Courier rate-card compliance checks.",
    "Zone, weight, failed-delivery, return, and surcharge analysis.",
    "Monthly founder dashboard.",
    "Dispute tracking.",
    "Quarterly courier performance review.",
]))

parts.append(heading("Tier 3: Logistics Governance — R29,000/month", 2))
parts.append(para("Best for brands spending more than R150,000 per month on courier costs, where a small percentage leak can become a six-figure annual problem."))
parts.append(bullets([
    "Everything in Tier 2.",
    "Weekly courier export audit.",
    "Monthly executive margin dashboard.",
    "Packaging governance by product or SKU category.",
    "Volumetric weight leakage tracking.",
    "Courier dispute workflow management.",
    "Courier SLA and charge accuracy scorecard.",
    "Quarterly rate-card and courier performance review.",
    "Negotiation support with courier providers.",
    "Board/founder-level logistics margin summary.",
]))

parts.append(page_break())
parts.append(heading("3. ROI Framing for Founders", 1))
parts.append(para("The retainer should not be framed as a promise to keep finding the same level of leakage every month. That is both commercially weak and operationally untrue if the system works."))
parts.append(para("Instead, frame the retainer as a margin-control function that prevents courier cost creep from returning."))

parts.append(heading("Founder Talking Points", 2))
parts.append(bullets([
    "The first audit recovers the backlog. Governance protects the future margin.",
    "Courier costs rarely explode overnight; they creep up through weight errors, packaging drift, surcharge changes, failed deliveries, returns, and rate-card exceptions.",
    "For less than the cost of one internal ops or finance hire, the founder gets a specialist logistics margin-control function.",
    "The goal is not to keep discovering disasters. The goal is to make sure courier cost disasters do not silently return.",
]))

parts.append(heading("Example ROI Calculation", 2))
parts.append(table([
    ["Metric", "Example Amount"],
    ["Monthly courier spend", "R150,000"],
    ["Conservative leakage assumption", "5%"],
    ["Monthly leakage risk", "R7,500"],
    ["Annual leakage risk", "R90,000"],
    ["Additional packaging inefficiency risk at 3%", "R54,000/year"],
    ["Combined annual margin risk", "R144,000/year"],
    ["Tier 2 governance investment", "R14,500/month"],
], [4200, 4200]))
parts.append(para("For larger clients, the economics become even stronger. A brand spending R300,000 per month on couriers has a 5% leakage exposure of R15,000 per month, or R180,000 per year, before packaging inefficiency, return charges, surcharge drift, or rate-card errors are included."))

parts.append(heading("Positioning Statement", 2))
parts.append(para("We recover historical courier overcharges, then run monthly logistics governance so your delivery costs do not quietly eat your margin as you scale.", bold=True, color="1F4E79"))

parts.append(page_break())
parts.append(heading("4. Performance Guarantee", 1))
parts.append(heading("No-Brainer Governance Guarantee", 2))
parts.append(para("If we do not identify, recover, or prevent logistics leakage equal to at least your monthly governance fee over any rolling 90-day period, we credit the shortfall against your next invoice.", bold=True, color="1F4E79"))
parts.append(para("The guarantee applies as long as we receive the required courier exports, invoices, rate cards, and packaging data on time.", bold=True, color="1F4E79"))

parts.append(heading("Why This Guarantee Works", 2))
parts.append(bullets([
    "It removes financial anxiety for the founder.",
    "It avoids exposing the agency to full cash refunds after the work has been performed.",
    "It forces the client to provide the data needed to succeed.",
    "It validates prevented leakage as legitimate value, not only recovered cash.",
]))

parts.append(heading("Month 1 Historical Audit Guarantee", 2))
parts.append(para("If we do not identify verified recoveries or savings worth at least 2x the audit fee, we waive the success fee and credit 50% of the audit fee toward your first month of governance.", bold=True, color="1F4E79"))
parts.append(para("For example, on a R32,000 audit, the target is at least R64,000 in verified value. If that threshold is not met, the client pays no success fee and receives a R16,000 credit toward governance."))

parts.append(page_break())
parts.append(heading("5. Recommended Flagship Offer", 1))
parts.append(para("To reduce sales complexity, lead with one flagship package and adjust only when the client's courier spend clearly falls below or above the target range."))

parts.append(table([
    ["Phase", "Price", "Includes"],
    ["Month 1: Historical Audit", "R32,000 once-off + 22.5% of verified recoveries and first-60-day packaging savings", "30–90 day historical audit; courier overcharge detection; weight/volumetric review; zone, fuel, surcharge, failed-delivery and return charge review; packaging inefficiency diagnosis; dispute pack; founder ROI report"],
    ["Month 2+: Governance", "R14,500/month or R29,000/month depending on courier spend", "Monthly or weekly audit cadence; anomaly review; dispute tracking; packaging governance; founder dashboard; quarterly courier performance and rate review"],
], [1900, 3300, 4700]))

parts.append(heading("Minimum Contract Terms", 2))
parts.append(bullets([
    "Historical audit paid upfront.",
    "3-month minimum governance term after the historical audit.",
    "Month-to-month thereafter with 30 days' notice.",
    "Retainer billed monthly in advance.",
    "Success fees billed monthly in arrears once value is verified.",
    "Client must provide courier exports, invoices, rate cards, packaging data, and dispute history within agreed timelines.",
]))

parts.append(heading("Final Strategic Position", 2))
parts.append(para("Do not sell 'we find mistakes.' Sell 'monthly logistics margin control for South African e-commerce founders.'", bold=True, color="1F4E79"))
parts.append(para("This positions the business as a durable governance partner rather than a once-off audit vendor, while still capturing the significant initial upside from historical recoveries."))

body = ''.join(parts)

document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
'''

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:color="BFBFBF"/></w:tblBorders></w:tblPr></w:style>
</w:styles>
'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'''

with ZipFile(OUT, "w", ZIP_DEFLATED) as docx:
    docx.writestr("[Content_Types].xml", content_types)
    docx.writestr("_rels/.rels", rels)
    docx.writestr("word/document.xml", document_xml)
    docx.writestr("word/styles.xml", styles_xml)
    docx.writestr("word/_rels/document.xml.rels", doc_rels)

print(OUT)
