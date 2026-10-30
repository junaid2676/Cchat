"""Builds the CChat Decision Log (.docx)."""
import sys
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

PURPLE = RGBColor(0x6D, 0x28, 0xD9)
DARK = RGBColor(0x1F, 0x1B, 0x2E)
MUTED = RGBColor(0x55, 0x55, 0x66)

OUT = sys.argv[1]

# ---------------------------------------------------------------- content
DECISIONS = [
    # ---------------- Product
    dict(area="Product", title="Build CChat: an AI receptionist for small businesses", date="26 Oct 2026", status="Decided",
         what="After practising with CloudPulse (website monitor) and Nestwise (a demo website), I chose a project that solves a real problem: small businesses miss phone calls and orders at busy times. CChat answers the phone, takes orders and hands them to the shop.",
         why="Recruiters for cloud roles want to see real-world problems solved end to end. This project needs almost every core skill: backend, databases, security, AI, telephony, CI/CD and AWS.",
         alt="Keep improving the CloudPulse monitor, or finish the Nestwise website. Both were kept as earlier learning projects rather than deleted.",
         learned="A portfolio project is stronger when it starts from a business problem, not from a technology.",
         line="I picked a problem real shops have, missed calls at peak times, and built the whole system around it."),
    dict(area="Product", title="Target customer: busy small shops with a till person", date="27 Oct 2026", status="Decided",
         what="CChat is for small UK shops (takeaways, cafés, grocers, parcel shops, retail) that get swamped with calls at peak times, where a dedicated till person usually takes orders.",
         why="A clear target customer makes every later decision easier. Orders go to the till first because that is how these shops already work.",
         alt="Targeting large restaurants or chains, which already have expensive phone systems.",
         learned="Product teams call this defining the target customer (a persona). Features are judged against what that person actually needs.",
         line="Every feature was checked against one customer: a small shop drowning in calls at 7pm."),
    dict(area="Product", title="Name: CChat", date="27 Oct 2026", status="Decided",
         what="Named the product CChat, with the tagline 'Your customers call, we chat'. An earlier shortlist included Counterline and Tillcall.",
         why="Short and easy to remember. The tagline makes it clear the product answers phone calls, not just text chat.",
         alt="Counterline (shop counter + phone line) and Tillcall.",
         learned="Check a name against existing products before using it. For a real business, also search the UK trademark register and Companies House.",
         line=""),
    dict(area="Product", title="Phone calls are the core channel", date="26 Oct 2026", status="Decided",
         what="Customers ring from landlines and mobiles, so voice is the main channel. The AI 'brain' is built and tested by typing first, then voice is added as early as Stage 6.",
         why="A landline caller cannot type. But testing by typing is instant and free, while every test call costs money and time, so the same brain is built first and then given a voice.",
         alt="Voice later in the roadmap (my original plan, which I challenged and changed).",
         learned="Separate how you build (testing order) from what the customer uses. One brain, many channels.",
         line="I designed one agent brain that serves both chat and phone, so voice became a channel rather than a separate system."),
    dict(area="Product", title="Businesses keep their own phone number", date="26 Oct 2026", status="Decided",
         what="Shops forward their existing number to a CChat number (Amazon Connect): always, only when busy or unanswered, or after hours.",
         why="No customer has to learn a new number, and the shop can start with overflow-only forwarding, which is an easy first step for a nervous owner.",
         alt="Porting the number into Amazon Connect (bigger commitment), or advertising a brand new number.",
         learned="How phone calls reach a cloud system: the public phone network, call forwarding, and DID numbers.",
         line="Shops keep their number and simply forward calls, so adoption needs zero change for their customers."),
    # ---------------- Call workflow
    dict(area="Workflow", title="Call workflow designed in five versions", date="27 Oct 2026", status="Done (v5 final)",
         what="I drew the end-to-end call flow and improved it over five versions. It added an opening-hours check, decision diamonds, a confirm-the-order step, a re-ask loop, escalation to a human, staff acceptance, optional marketing consent and a warm goodbye.",
         why="Designing before coding is cheaper: a mistake in a diagram costs minutes, while the same mistake in code costs days.",
         alt="Start coding straight away and discover the flow later.",
         learned="Flowchart conventions (decision diamonds, labelled branches), loops (like a Python while loop), and knowing when to stop polishing a design.",
         line="I iterated the call flow five times before writing code; the read-back loop and escalation path came from that process."),
    dict(area="Workflow", title="Read the order back, and escalate if it is still wrong", date="27 Oct 2026", status="Decided",
         what="The AI reads the order and total back. If the customer corrects it, the AI re-asks once; if it is still not right, the call is passed to a human.",
         why="Speech recognition makes mistakes ('fifteen' heard as 'fifty'). A loop that could repeat forever traps callers, so escalation gives them a way out.",
         alt="Send orders straight to the shop without confirming (my first version).",
         learned="Escalation is a standard pattern in support systems: automate, but always offer a human route.",
         line="Wrong orders are prevented by a read-back loop with an escalation path, so no caller gets stuck with the bot."),
    dict(area="Workflow", title="Never take card details over the phone", date="26 Oct 2026", status="Decided",
         what="The AI sends a secure payment link by SMS instead of asking for card numbers. Payment timing (pay first, pay on collection, customer chooses) is an owner setting.",
         why="Card data is covered by PCI DSS, a strict security standard. Never touching card numbers removes a huge security and legal burden.",
         alt="Let the AI read card numbers over the phone.",
         learned="PCI DSS, and reducing risk by designing data out of the system instead of protecting it.",
         line="The safest card data is data you never hold, so payments go through a hosted payment link."),
    # ---------------- Scope
    dict(area="Scope", title="CChat hands orders off; it does not track fulfilment", date="28 Oct 2026", status="Decided",
         what="CChat's job ends when the order reaches the shop and payment is handled. It tracks the handoff (sent, or failed with automatic retry and an alert) and payment, but not 'preparing', 'ready' or 'completed'.",
         why="Shops already have a till or POS for that. Duplicating it would create two sources of truth, and staff would stop updating one of them. The failure that really matters is an order that never reaches the shop.",
         alt="Full order status tracking (an earlier suggestion I challenged).",
         learned="Scope control, the 'system of record' idea, and designing for the most damaging failure (a silent handoff failure).",
         line="I deliberately kept fulfilment out of scope so the POS stays the single source of truth, and focused on making handoff failures impossible to miss."),
    dict(area="Scope", title="No stock tracking", date="28 Oct 2026", status="Decided",
         what="Items have a simple Available on/off switch instead of stock counts.",
         why="Stock numbers are only correct if every in-shop sale updates them, which the till does, not CChat. A wrong stock count would make the AI tell callers something false. Real stock levels could come later from a POS integration.",
         alt="Owners typing stock counts into CChat (it crept into the first Figma design).",
         learned="The same scope rule applied again: only store what CChat can keep accurate.",
         line=""),
    dict(area="Scope", title="Generic for any small business", date="28 Oct 2026", status="Decided",
         what="The app says 'Items', never 'Menu'. Items have option groups (size, extras, delivery speed) and business-specific fields chosen by business type (allergens for food, weight limits for parcels).",
         why="Food is only one example; the same product must work for a café, a grocer, a post office or a retail shop.",
         alt="A restaurant-only app.",
         learned="Designing a flexible data model: shared core fields plus type-specific details.",
         line="One data model serves a pizza, a pint of milk and a next-day parcel."),
    dict(area="Scope", title="The AI never guesses about allergens", date="28 Oct 2026", status="Decided",
         what="If allergen information is not filled in, the AI passes allergy questions to staff. 'Contains none of the 14' is a separate tick from 'not filled in'.",
         why="UK food law requires accurate information on 14 major allergens, and a wrong answer can put someone in hospital.",
         alt="Treat an empty allergen list as 'none'.",
         learned="Safety-critical data needs an explicit 'unknown' state.",
         line="Empty is not the same as none: the AI only says an item is allergen-free when the owner has actively confirmed it."),
    # ---------------- Privacy
    dict(area="Privacy", title="Three consent states; feedback is separate from marketing", date="28 Oct 2026", status="Decided",
         what="Customers are Opted in, Feedback only, or Opted out. Feedback links are allowed for any customer who ordered; offers only for opted-in customers. Every offer includes 'Reply STOP to opt out'. Each consent is stored with how and when it was given.",
         why="UK GDPR and PECR control marketing messages. Building the rules into the app means owners cannot break the law by accident.",
         alt="A single yes/no flag, which cannot tell 'never asked' from 'said stop'.",
         learned="PECR, UK GDPR, the ICO, consent records as proof, and the right to erasure (a 'delete customer data' button).",
         line="Compliance is built into the interface: the offer button is disabled unless consent exists, and every consent has an audit record."),
    dict(area="Privacy", title="Never ask for marketing consent inside a feedback message", date="28 Oct 2026", status="Decided",
         what="Opt-in happens only on the call or on the receipt page (an unticked checkbox), never in a message sent to someone who has not opted in.",
         why="Under PECR, a message asking for marketing consent counts as marketing itself. The ICO has fined large companies such as Honda and Flybe for exactly this.",
         alt="My first idea: add an opt-in request to the feedback text.",
         learned="To check a sensible-looking idea against the actual regulation.",
         line=""),
    # ---------------- Design
    dict(area="Design", title="One responsive web app instead of native apps", date="27 Oct 2026", status="Decided",
         what="A single HTML/CSS web app that adapts to phone, tablet (till screen) and PC, which can later be installable as a PWA.",
         why="One codebase, no app store approvals, and it reuses my HTML/CSS skills. That is the right trade-off for a team of one.",
         alt="Separate iPhone and Android apps.",
         learned="Responsive, mobile-first design and Progressive Web Apps.",
         line=""),
    dict(area="Design", title="Workflow, then wireframes, then high-fidelity designs", date="28 Oct 2026", status="Done",
         what="I designed the call flow first, then low-detail wireframes, then high-fidelity screens using the Figma AI agent from precise written specifications.",
         why="Deciding the data and the logic myself, then letting AI speed up the visuals. Every screen element was checked against where its data will come from.",
         alt="Jump straight to polished visuals.",
         learned="Organise screens around things (orders, items), not fields (price, allergens). Bottom tab bars for phones. Design every state (empty, error, no permission).",
         line="Each screen was reviewed against the data model, so the designs also became the database design."),
    # ---------------- Security
    dict(area="Security", title="Roles and least privilege", date="27 Oct 2026", status="Decided",
         what="Owners can change prices and see revenue and customers; till staff can only see orders and calls and switch item availability. The till tablet stays logged in with the limited staff role, and the owner can revoke it.",
         why="Staff don't need the power to change every price. Least privilege limits the damage from mistakes or a stolen device.",
         alt="One shared login for everyone.",
         learned="Least privilege, the same principle behind AWS IAM.",
         line=""),
    dict(area="Security", title="Invite-first onboarding, sign-up with an approval gate", date="28 Oct 2026", status="Decided",
         what="New businesses can sign up and explore, but calls and SMS only switch on after manual approval. Invited owners receive a one-time setup link that expires in 48 hours, not a temporary password.",
         why="It stops fake accounts using CChat to send spam. A one-time link cannot sit in an inbox as a working password.",
         alt="Open self sign-up with everything enabled immediately, or emailing temporary passwords.",
         learned="Concierge onboarding, and why expiring single-use links are safer than temporary passwords.",
         line=""),
    dict(area="Security", title="Business ID is internal, random and never typed", date="28 Oct 2026", status="Decided",
         what="Every database row carries a business_id (multi-tenancy), but users log in with email only. IDs are random UUIDs, not counting numbers.",
         why="An ID is not a secret, so typing it adds friction without security. Counting IDs invite IDOR attacks (changing 41 to 42 to see another shop's data).",
         alt="A mandatory business ID field at login (my first idea).",
         learned="Multi-tenancy, UUIDs and the IDOR vulnerability.",
         line=""),
    dict(area="Security", title="Two-factor login with an authenticator app or passkey", date="28 Oct 2026", status="Decided",
         what="Owners use an authenticator app or a passkey as the second step, with 10 recovery codes. SMS codes are a backup only, and email codes are not used.",
         why="The app holds customer phone numbers and revenue, so a stolen password alone must not be enough. SMS is open to SIM-swap fraud.",
         alt="Password only; SMS or email codes as the main method.",
         learned="MFA methods and their weaknesses, passkeys, and recovery codes.",
         line=""),
    dict(area="Security", title="Login messages that reveal nothing", date="28 Oct 2026", status="Decided",
         what="'Email or password is incorrect' and 'If an account exists for that email, we've sent a link' are shown whether or not the email exists.",
         why="This prevents account enumeration: attackers discovering which shops use CChat.",
         alt="Specific errors such as 'No account with this email'.",
         learned="Account enumeration attacks.",
         line=""),
    dict(area="Security", title="Use Amazon Cognito rather than building login", date="28 Oct 2026", status="Decided",
         what="Sign-up, login, 2FA, passkeys and password resets will be handled by Amazon Cognito.",
         why="Writing your own authentication is a classic security mistake. Cognito is battle-tested, and configuring it teaches how authentication works.",
         alt="A custom-built login system.",
         learned="Use managed, proven services for security-critical parts.",
         line=""),
    # ---------------- Cloud
    dict(area="Cloud", title="AWS London region (eu-west-2)", date="27 Oct 2026", status="Decided",
         what="All resources run in Europe (London).",
         why="The customers are UK businesses. Keeping data in the UK supports UK GDPR and lowers latency for callers.",
         alt="A US region (sometimes cheaper).",
         learned="Data residency.",
         line=""),
    dict(area="Cloud", title="Walking skeleton: deploy early, in stages", date="28 Oct 2026", status="Decided",
         what="Build the thinnest version that goes all the way through (code, tests, container, pipeline, AWS, live URL), then add features stage by stage: skeleton, real login (Cognito), data (RDS), Terraform, agent brain (Bedrock), phone (Amazon Connect).",
         why="Deployment always brings surprises. It is far cheaper to meet them while the app is tiny, and CI/CD is demonstrated from day one.",
         alt="Build everything locally and deploy at the end.",
         learned="Walking skeleton, CI/CD, and incremental delivery.",
         line="I had CI/CD and a live AWS deployment before the app had features, so every change after that was deployed safely and automatically."),
    dict(area="Cloud", title="Stage 1 stack: FastAPI, Jinja2, Docker, GitHub Actions, EC2", date="28 Oct 2026", status="Planned",
         what="Python FastAPI serves HTML pages through Jinja2 templates, packaged with Docker, tested and deployed by GitHub Actions to an EC2 server.",
         why="It keeps me in HTML, CSS and Python, Docker makes the app run the same everywhere, and starting on EC2 teaches the fundamentals (servers, networking, firewalls) before moving to managed containers.",
         alt="A JavaScript framework for the frontend; managed containers (ECS) from the start.",
         learned="To be filled in during Stage 1.",
         line=""),
    dict(area="Cloud", title="Lock down the AWS account before creating anything", date="29 Oct 2026", status="Done",
         what="MFA on the root user; a recurring $10 monthly budget (unblended costs, alerts at 50% and 80% actual and 100% forecasted; credits excluded once Cost Explorer data appears); and a personal IAM admin user with MFA for daily work. Root is signed out and kept for rare account-level tasks only.",
         why="The root user can do anything, including deleting the account. Budgets are only alarms (they never charge or stop anything), so they must count real usage, not usage after credits, or they would never fire.",
         alt="Working as root with no budget alerts.",
         learned="Root vs IAM users, MFA, AWS Budgets, unblended costs, and that Cost Explorer data can take about 24 hours to appear on a new account.",
         line="Before creating any resource I secured root with MFA, created a separate admin user with MFA and set forecast-based budget alerts."),
    dict(area="Cloud", title="IAM user instead of IAM Identity Center (for now)", date="28 Oct 2026", status="Decided",
         what="My daily admin login is an IAM user in an 'Admins' group with MFA, console access only and no access keys.",
         why="Identity Center needs AWS Organizations to grant account access, and enrolling a Free plan account in Organizations converts it to the Paid plan and removes the Free Tier credits. For a single-account project, an IAM user with MFA is equally secure.",
         alt="IAM Identity Center with AWS Organizations (AWS's recommendation for multi-account setups). I'll revisit it after upgrading to the Paid plan.",
         learned="To read the terms before enabling a service: an innocent-looking setting can change billing. Permissions belong on groups, not on individual users.",
         line="I checked the Free plan terms before enabling Identity Center, found it would forfeit the credits, and chose an IAM user with MFA instead."),
    dict(area="Cloud", title="Demo data only while on the Free plan", date="28 Oct 2026", status="Decided",
         what="Until the account is upgraded, CChat stores only fictional test data: no real customer names or phone numbers.",
         why="AWS's Free plan terms say Free plan accounts are for evaluation and should not be used to process sensitive data.",
         alt="",
         learned="Account plans carry usage terms as well as prices.",
         line=""),
    dict(area="Cloud", title="Free plan now, upgrade before it ends", date="28 Oct 2026", status="Noted",
         what="The account is on the AWS Free plan (credits, about 6 months, ending around mid-April 2027). It will be upgraded to the Paid plan before then, keeping any remaining credits.",
         why="The heavier services (Bedrock, Amazon Connect) arrive in later stages, and the Free plan closes when its time ends.",
         alt="",
         learned="AWS Free plan vs Paid plan.",
         line=""),
]

GLOSSARY = [
    ("Multi-tenant", "One system serves many businesses, and each one's data is walled off (every row carries a business_id)."),
    ("Intent", "What a caller wants: an order, a question, a human or a message."),
    ("Escalation", "Handing over to a human when automation can't resolve something."),
    ("Scope", "What a product deliberately does and does not do."),
    ("System of record", "The one place that holds the official version of some data (for fulfilment: the shop's POS)."),
    ("Audit trail", "A timestamped record of what happened, used as proof (for example, order timeline, consent record)."),
    ("Least privilege", "Give each person or system only the access it needs."),
    ("PCI DSS", "Security standard for handling card data."),
    ("UK GDPR / PECR", "UK data protection law / UK rules for marketing calls, texts and emails."),
    ("IDOR", "Insecure Direct Object Reference: accessing someone else's data by changing an ID."),
    ("Account enumeration", "Discovering which accounts exist through revealing error messages."),
    ("MFA / 2FA", "A second login step beyond the password."),
    ("Walking skeleton", "The thinnest end-to-end version of a system, deployed early."),
    ("CI/CD", "Automatically testing (CI) and deploying (CD) every change."),
    ("MVP", "Minimum viable product: the smallest version worth using."),
]

# ---------------------------------------------------------------- helpers
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)  # A4
for side in ("left_margin", "right_margin"):
    setattr(sec, side, Cm(2.2))
sec.top_margin = sec.bottom_margin = Cm(2)

normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(11)
normal.font.color.rgb = DARK
normal.paragraph_format.space_after = Pt(4)
for name, size in (("Title", 26), ("Heading 1", 16), ("Heading 2", 13)):
    st = doc.styles[name]
    st.font.name = "Calibri"
    st.font.size = Pt(size)
    st.font.color.rgb = PURPLE
    st.font.bold = True


def shade(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def labelled(label, text, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(label + ": ")
    r.bold = True
    r.font.color.rgb = PURPLE
    body = p.add_run(text)
    body.italic = italic
    return p


def muted(text, size=10):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.color.rgb = MUTED
    return p


# ---------------------------------------------------------------- document
doc.add_paragraph("CChat: Decision Log", style="Title")
muted("What I decided, why, and what I learned along the way", 12)
muted("Author: Junaid  ·  Started: 26 October 2026  ·  Last updated: 29 October 2026")

doc.add_heading("About this document", level=1)
doc.add_paragraph(
    "CChat is an AI receptionist for small UK businesses: it answers the shop's phone, takes orders, "
    "sends them to the till and handles payment links, so busy shops stop missing calls. I am building it "
    "end to end, from design to a live deployment on AWS with CI/CD, as my cloud engineering portfolio project.")
doc.add_paragraph(
    "This log records every important decision in the same format: what I decided, why, what else I "
    "considered, and what I learned. Where useful, an interview line sums the decision up in one sentence. "
    "It is a living document, updated at every stage of the build.")

doc.add_heading("Decisions at a glance", level=1)
table = doc.add_table(rows=1, cols=4)
table.style = "Table Grid"
table.alignment = WD_TABLE_ALIGNMENT.CENTER
widths = (Cm(1.0), Cm(2.4), Cm(10.2), Cm(3.0))
for i, head in enumerate(("#", "Area", "Decision", "Status")):
    cell = table.rows[0].cells[i]
    cell.text = ""
    run = cell.paragraphs[0].add_run(head)
    run.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    shade(cell, "6D28D9")
for n, d in enumerate(DECISIONS, 1):
    row = table.add_row().cells
    for i, val in enumerate((str(n), d["area"], d["title"], d["status"])):
        row[i].text = val
        row[i].paragraphs[0].runs[0].font.size = Pt(9.5)
        if n % 2 == 0:
            shade(row[i], "F3EEFF")
for row in table.rows:
    for i, w in enumerate(widths):
        row.cells[i].width = w

doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
doc.add_heading("The decisions", level=1)

current_area = None
for n, d in enumerate(DECISIONS, 1):
    if d["area"] != current_area:
        current_area = d["area"]
        p = doc.add_paragraph()
        r = p.add_run(current_area.upper())
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = MUTED
        p.paragraph_format.space_before = Pt(10)
    doc.add_heading(f"{n}. {d['title']}", level=2)
    muted(f"{d['date']}  ·  Status: {d['status']}", 9.5)
    labelled("What", d["what"])
    labelled("Why", d["why"])
    if d["alt"]:
        labelled("Alternatives considered", d["alt"])
    labelled("What I learned", d["learned"])
    if d["line"]:
        labelled("Interview line", "“" + d["line"] + "”", italic=True)

doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
doc.add_heading("Concepts I've learned", level=1)
gt = doc.add_table(rows=0, cols=2)
gt.style = "Table Grid"
for term, meaning in GLOSSARY:
    cells = gt.add_row().cells
    cells[0].text = term
    cells[0].paragraphs[0].runs[0].bold = True
    cells[1].text = meaning
    shade(cells[0], "F3EEFF")
for row in gt.rows:
    row.cells[0].width, row.cells[1].width = Cm(4.2), Cm(12.4)

doc.add_heading("Where I am now", level=1)
for text in (
    "Design: call workflow (v5) and high-fidelity screens for login, 2FA, business switcher, Home, Orders and Items are done.",
    "AWS: account locked down (root MFA, budget alerts, IAM admin user with MFA).",
    "Next: Stage 1, the walking skeleton (FastAPI + Docker + GitHub Actions + EC2).",
):
    doc.add_paragraph(text, style="List Bullet")

doc.save(OUT)
print("saved", OUT, "decisions:", len(DECISIONS))
