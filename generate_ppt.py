from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()

# Slide 1: Title
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "Blockchain-Based Order Tracking for Cold Chain Logistics"
subtitle.text = "BCSE329L: Blockchain And Distributed Ledger Technology\nWinter Semester 2026-2027"

# Helper for content slides
def add_content_slide(prs, title_text, content_list):
    bullet_slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes
    title_shape = shapes.title
    body_shape = shapes.placeholders[1]
    title_shape.text = title_text
    
    tf = body_shape.text_frame
    for i, item in enumerate(content_list):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(20)

# Slide 2: Introduction
add_content_slide(prs, "1. Introduction & Background", [
    "Cold chain logistics transport temperature-sensitive goods (e.g., vaccines, perishables).",
    "Maintaining specific temperature ranges (2°C to 8°C) is critical to prevent spoilage.",
    "Blockchain provides a decentralized, immutable ledger for supply chains.",
    "Smart contracts automate agreements and enforce conditions transparently.",
    "Integrating IoT sensors allows real-time environmental monitoring."
])

# Slide 3: Problem Statement
add_content_slide(prs, "2. Problem Statement", [
    "Traditional systems are centralized, creating data silos.",
    "It is difficult to determine accountability when temperature excursions occur.",
    "Data can be tampered with or suppressed by malicious intermediaries.",
    "A decentralized, automated tracking system is needed to log IoT data securely.",
    "Automated smart contracts can immediately flag compromised shipments."
])

# Slide 4: Literature Review
add_content_slide(prs, "3. Literature Review", [
    "Foundational Papers:",
    " - Nakamoto (2008): Introduced decentralized ledgers (Bitcoin).",
    " - Wood (2014): Introduced smart contracts (Ethereum).",
    " - Christidis & Devetsikiotis (2016): Bridging IoT and Blockchain.",
    "Recent Developments:",
    " - Focus on supply chain cost reduction and transparency (Kshetri, 2021).",
    " - Practical traceability in food/pharma cold chains (Wang et al., 2023; Kumar et al., 2024).",
    " - Identification of constraints and implementation barriers (Rejeb et al., 2024)."
])

# Slide 5: Real-World Problem & Gaps
add_content_slide(prs, "4. Real-World Use Case & Research Gap", [
    "Use Case: Tracking pharmaceutical shipments (2°C - 8°C).",
    "Student 1 Gap (Architecture): High transaction costs and latency when logging frequent IoT data on a public blockchain.",
    "Student 2 Gap (Smart Contracts): The Oracle Problem – ensuring hardware sensors provide authentic, untampered data before it hits the chain.",
    "Student 3 Gap (Adoption): Lack of user-friendly interfaces to bridge Web3 with traditional enterprise ERP systems."
])

# Slide 6: Proposed Solution
add_content_slide(prs, "5. Proposed Solution", [
    "Decentralized Application (dApp): Built on EVM-compatible networks.",
    "Order Creation: Sellers initialize shipments with defined parameters.",
    "IoT Logging: Real-time sensor data is pushed via APIs to the smart contract.",
    "Automated Validation: Contract executes ConditionValid = (Temp ≥ 2°C) ∧ (Temp ≤ 8°C).",
    "Immediate Flagging: Any breach automatically updates status to 'Condition Violated'."
])

# Slide 7: References
add_content_slide(prs, "References", [
    "Christidis, K., & Devetsikiotis, M. (2016). IEEE Access.",
    "Kshetri, N. (2021). International Journal of Information Management.",
    "Nakamoto, S. (2008). Decentralized Business Review.",
    "Rejeb, A., et al. (2024). International Journal of Logistics Management.",
    "Wood, G. (2014). Ethereum project yellow paper."
])

prs.save("Presentation.pptx")
print("Presentation generated successfully.")
