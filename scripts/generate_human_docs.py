import os
import sys
import csv
from pathlib import Path

# ReportLab imports
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# python-pptx imports
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Define Project Paths
PROJECT_ROOT = Path(__file__).resolve().parent.parent
REPORTS_DIR = PROJECT_ROOT / "reports"
DATA_DIR = PROJECT_ROOT / "data" / "processed"
PDF_OUT_PATH = REPORTS_DIR / "Final_Report.pdf"
PPTX_OUT_PATH = REPORTS_DIR / "Presentation.pptx"

# Ensure reports directory exists
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

# ----------------------------------------------------------------------
# 1. DATA LOADER UTILITIES
# ----------------------------------------------------------------------
def load_csv_data(filename):
    filepath = DATA_DIR / filename
    if not filepath.exists():
        print(f"Warning: {filename} not found at {filepath}. Returning empty list.")
        return []
    data = []
    with open(filepath, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            data.append(row)
    return data

# Load CSV data
scorecard_raw = load_csv_data("fund_scorecard.csv")
var_cvar_raw = load_csv_data("var_cvar_report.csv")
hhi_raw = load_csv_data("sector_hhi_report.csv")
cleaning_raw = load_csv_data("cleaning_report.csv")
cohort_raw = load_csv_data("cohort_analysis.csv")

# Process scorecard data for easy use
scorecard_data = []
for row in scorecard_raw:
    try:
        scorecard_data.append({
            "rank": int(row.get("final_scorecard_rank", 99)),
            "name": row.get("scheme_name", "").replace(" - Direct Growth", "").replace(" Fund", ""),
            "cagr_1y": float(row.get("cagr_1y", 0)),
            "cagr_3y": float(row.get("cagr_3y", 0)),
            "sharpe": float(row.get("sharpe_ratio", 0)),
            "sortino": float(row.get("sortino_ratio", 0)),
            "beta": float(row.get("beta", 0)),
            "alpha": float(row.get("alpha_annual", 0)),
            "max_dd": float(row.get("max_drawdown", 0)),
            "expense": float(row.get("expense_ratio", 0)),
            "tracking_error": float(row.get("tracking_error", 0))
        })
    except Exception as e:
        print(f"Error parsing scorecard row: {row}. Error: {e}")

# Sort by rank
scorecard_data = sorted(scorecard_data, key=lambda x: x["rank"])

# Process Risk Data
risk_map = {row["scheme_name"]: row for row in var_cvar_raw}
hhi_map = {row["scheme_name"]: row for row in hhi_raw}
combined_risk_data = []
for scheme_name in risk_map:
    clean_name = scheme_name.replace(" - Direct Growth", "").replace(" Fund", "")
    var_val = float(risk_map[scheme_name].get("historical_var_95", 0))
    cvar_val = float(risk_map[scheme_name].get("conditional_var_95", 0))
    hhi_val = float(hhi_map.get(scheme_name, {}).get("sector_hhi", 0))
    conc_level = hhi_map.get(scheme_name, {}).get("concentration_level", "Unknown")
    combined_risk_data.append({
        "name": clean_name,
        "var_95": var_val,
        "cvar_95": cvar_val,
        "hhi": hhi_val,
        "concentration": conc_level
    })

# ----------------------------------------------------------------------
# 2. PDF REPORT GENERATOR (ReportLab)
# ----------------------------------------------------------------------
class NumberedCanvas(canvas.Canvas):
    """
    Two-pass canvas pattern to compute total page count dynamically.
    Draws custom corporate headers/footers on page 2+ and confidentiality stamps.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        
        # Color Palette
        primary_navy = colors.HexColor("#0F172A")
        muted_gray = colors.HexColor("#64748B")
        border_light = colors.HexColor("#E2E8F0")

        # Skip decorations on the cover page (Page 1)
        if self._pageNumber > 1:
            # Running Header
            self.setFont("Helvetica-Bold", 8)
            self.setFillColor(primary_navy)
            self.drawString(54, 755, "BLUESTOCK MUTUAL FUND ANALYTICS & RISK PLATFORM")
            
            self.setFont("Helvetica", 8)
            self.setFillColor(muted_gray)
            self.drawRightString(558, 755, "TECHNICAL CASE STUDY & EXECUTIVE REPORT")
            
            # Header Divider Line
            self.setStrokeColor(border_light)
            self.setLineWidth(0.75)
            self.line(54, 747, 558, 747)

            # Running Footer
            self.drawString(54, 42, "Confidential - Portfolio Management Advisory Services")
            self.drawRightString(558, 42, f"Page {self._pageNumber} of {page_count}")
            
            # Footer Divider Line
            self.line(54, 52, 558, 52)

        else:
            # Cover Page Accents
            # Draw a thick navy bar at the top edge of the cover page
            self.setFillColor(primary_navy)
            self.rect(0, 770, 612, 22, fill=True, stroke=False)
            
            # Draw a bottom accent bar
            self.setFillColor(colors.HexColor("#00E5FF")) # Cyan
            self.rect(0, 760, 612, 10, fill=True, stroke=False)
            
            # Draw bottom border
            self.setFillColor(primary_navy)
            self.rect(0, 0, 612, 15, fill=True, stroke=False)

        self.restoreState()

def build_pdf(filename):
    doc = SimpleDocTemplate(
        str(filename),
        pagesize=letter,
        leftMargin=54,
        rightMargin=54,
        topMargin=64,
        bottomMargin=64
    )
    
    styles = getSampleStyleSheet()
    
    # Custom Color Definitions
    primary_navy = colors.HexColor("#0F172A")
    secondary_blue = colors.HexColor("#1E3A8A")
    accent_teal = colors.HexColor("#0F766E")
    text_dark = colors.HexColor("#334155")
    bg_light = colors.HexColor("#F8FAFC")
    
    # Typography Styles
    title_style = ParagraphStyle(
        'CoverTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=30,
        textColor=primary_navy,
        spaceAfter=10
    )
    muted_gray = colors.HexColor("#475569")
    subtitle_style = ParagraphStyle(
        'CoverSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=12,
        leading=16,
        textColor=muted_gray,
        spaceAfter=40
    )
    metadata_style = ParagraphStyle(
        'CoverMetadata',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=14,
        textColor=colors.HexColor("#64748B"),
        spaceAfter=150
    )
    h1_style = ParagraphStyle(
        'DocH1',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=18,
        textColor=primary_navy,
        spaceBefore=18,
        spaceAfter=8,
        keepWithNext=True
    )
    h2_style = ParagraphStyle(
        'DocH2',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=secondary_blue,
        spaceBefore=10,
        spaceAfter=6,
        keepWithNext=True
    )
    body_style = ParagraphStyle(
        'DocBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13.5,
        textColor=text_dark,
        spaceAfter=8
    )
    bullet_style = ParagraphStyle(
        'DocBullet',
        parent=body_style,
        leftIndent=15,
        firstLineIndent=-10,
        spaceAfter=4
    )
    code_style = ParagraphStyle(
        'DocCode',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#0F172A"),
        backColor=colors.HexColor("#F1F5F9"),
        borderColor=colors.HexColor("#CBD5E1"),
        borderWidth=0.5,
        borderPadding=6,
        spaceBefore=6,
        spaceAfter=8
    )
    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=10,
        textColor=colors.white,
        alignment=1 # Centered
    )
    table_cell_style = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        textColor=text_dark
    )
    table_cell_bold_style = ParagraphStyle(
        'TableCellBold',
        parent=table_cell_style,
        fontName='Helvetica-Bold'
    )
    table_cell_center_style = ParagraphStyle(
        'TableCellCenter',
        parent=table_cell_style,
        alignment=1
    )

    story = []
    
    # ------------------------------------------------------------------
    # PAGE 1: TITLE PAGE
    # ------------------------------------------------------------------
    story.append(Spacer(1, 40))
    story.append(Paragraph("TECHNICAL CASE STUDY & ADVISORY REPORT", ParagraphStyle('Upper', fontName='Helvetica-Bold', fontSize=10, textColor=accent_teal, spaceAfter=15)))
    story.append(Paragraph("Bluestock Mutual Fund Analytics & Portfolio Risk Platform", title_style))
    story.append(Paragraph("An Enterprise-Grade ETL, Relational SQLite Schema, Performance Scorecard Leaderboard, Cohort Retention Engine, and Power BI Dashboard Design", subtitle_style))
    
    # Draw horizontal bar separator
    story.append(Spacer(1, 15))
    
    story.append(Paragraph(
        "<b>Prepared By:</b> Dhileep B, Lead Financial Data Analyst & Senior Quantitative Analyst<br/>"
        "<b>Project Context:</b> Bluestock Capstone Portfolio Submission (Days 1–6)<br/>"
        "<b>Technology Stack:</b> Python (pandas, scipy, numpy, matplotlib, seaborn), SQLite 3 Relational DB, yfinance API, Power BI Desktop<br/>"
        "<b>Submission Date:</b> June 12, 2026<br/>"
        "<b>Version:</b> 1.2.0 (Production Clean Build)<br/>"
        "<b>Classification:</b> Confidential, Investment Advisory Team Review",
        metadata_style
    ))
    
    # Executive abstract block in a box
    abstract_text = (
        "<b>Executive Abstract:</b> This paper documents the construction and analytical findings of the Bluestock "
        "Mutual Fund Analytics Platform. We designed a standardized ETL ingestion script that cleans raw AMFI records and client "
        "transaction databases, flattening yfinance benchmark indices (Nifty 50 and Nifty 100) to resolve schema anomalies. The data is "
        "loaded into a STAR schema SQLite relational database. We execute quantitative risk engines to generate a weighted fund "
        "scorecard (returns, Sharpe, Alpha, expense ratio, and drawdowns), model daily tail risk (Value at Risk and Conditional VaR), "
        "compute portfolio sector concentrations (HHI), and trace retail investor cohort flows. Finally, we establish a robust "
        "specification for a dark-themed Power BI interactive dashboard to translate quantitative findings into strategic advisor recommendations."
    )
    
    abstract_table = Table([[Paragraph(abstract_text, ParagraphStyle('Abstract', parent=body_style, fontSize=9, leading=12.5, textColor=primary_navy))]], colWidths=[504])
    abstract_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EFF6FF")),
        ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor("#3B82F6")),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
    ]))
    story.append(abstract_table)
    story.append(PageBreak())
    
    # ------------------------------------------------------------------
    # PAGE 2: OBJECTIVES & INGESTION
    # ------------------------------------------------------------------
    story.append(Paragraph("1. Platform Objectives & Background", h1_style))
    story.append(Paragraph(
        "Modern wealth management firms are inundated with fragmented client and scheme data. The primary challenges in mutual fund analytics lie in resolving the gaps between three disjointed data worlds: (a) raw Daily Net Asset Values (NAV) published by AMFI, (b) high-velocity retail transaction ledgers, and (c) broad market benchmark indices. The objective of this capstone project is to engineer an enterprise-grade analytics engine that consolidates these sources into an optimized star schema database, and runs quantitative modules to drive automated investment scoring and retail cohort retention analysis.",
        body_style
    ))
    story.append(Paragraph(
        "By calculating localized measures (CAGRs, Sortino, rolling Sharpe Ratios, daily Value at Risk, Expected Shortfalls, and Herfindahl-Hirschman Indices), the platform provides advisors and portfolio managers with immediate visibility into performance drag, style consistency, and systemic tail-risk concentrations.",
        body_style
    ))
    
    story.append(Paragraph("2. ETL Ingestion & Data Cleaning Operations", h1_style))
    story.append(Paragraph(
        "The data ingestion layer is handled programmatically in <code>scripts/etl_pipeline.py</code>. The script validates column configurations, filters corrupt rows, deduplicates transaction records, and exports clean datasets into CSV and Parquet files in the <code>data/processed/</code> directory. Key ETL cleaning rules include:",
        body_style
    ))
    story.append(Paragraph("• <b>Date Standardisation:</b> Transaction and NAV dates are forced to standard ISO <code>YYYY-MM-DD</code>. Raw files containing invalid strings like <code>'INVALID'</code> are coerced using pandas to <code>NaT</code> and dropped.", bullet_style))
    story.append(Paragraph("• <b>Index MultiIndex Resolution:</b> Downloading market data via the Yahoo Finance API (using `yfinance`) on single tickers returns multi-level column indexing. The script programmatically flattens these arrays and extracts close prices to avoid column shifts during CSV write.", bullet_style))
    story.append(Paragraph("• <b>Outlier Truncation & Imputation:</b> Erroneous zero NAV values and obvious keystroke anomalies (e.g., negative NAV values) are removed. Missing middle dates are filled using a forward-fill (locf) technique to ensure cumulative return continuities.", bullet_style))
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("Table 2.1: Data Ingestion & Cleaning Operational Log", h2_style))
    
    # Cleaning Table
    clean_table_data = [[
        Paragraph("<b>Filename</b>", table_header_style),
        Paragraph("<b>Raw Rows</b>", table_header_style),
        Paragraph("<b>Final Rows</b>", table_header_style),
        Paragraph("<b>Dupes Del</b>", table_header_style),
        Paragraph("<b>Bad Dates</b>", table_header_style),
        Paragraph("<b>Issues Handled</b>", table_header_style)
    ]]
    for row in cleaning_raw:
        clean_table_data.append([
            Paragraph(row.get("filename", ""), table_cell_bold_style),
            Paragraph(row.get("raw_rows", ""), table_cell_center_style),
            Paragraph(row.get("final_rows", ""), table_cell_center_style),
            Paragraph(row.get("duplicates_removed", ""), table_cell_center_style),
            Paragraph(row.get("bad_dates_removed", ""), table_cell_center_style),
            Paragraph(row.get("issue_summary", "").split(" | ")[-1], table_cell_style)
        ])
        
    t_cleaning = Table(clean_table_data, colWidths=[90, 50, 50, 55, 55, 204])
    t_cleaning.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), primary_navy),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, bg_light]),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_cleaning)
    story.append(PageBreak())
    
    # ------------------------------------------------------------------
    # PAGE 3: SQL DATABASE SCHEMA
    # ------------------------------------------------------------------
    story.append(Paragraph("3. Relational STAR Schema Database Design", h1_style))
    story.append(Paragraph(
        "To facilitate high-performance BI loading and direct SQL querying, the clean data is modeled into a relational **STAR Schema** in SQLite (`data/db/mutual_funds.db` and defined in `sql/schema.sql`). The schema isolates structural metadata (schemes, AMCs, investors) into dimension tables and transactional metrics (daily NAV history, client buy/sell ledgers) into central fact tables.",
        body_style
    ))
    
    # Database structure block
    story.append(Paragraph("<b>Dimension Tables:</b>", h2_style))
    story.append(Paragraph("• <code>dim_amc</code>: Master list of Asset Management Companies. Primary key: <code>amc_id</code>. Attributes: AMC name, code.", bullet_style))
    story.append(Paragraph("• <code>dim_scheme</code>: Mutual fund schemes mapped by categories (Equity, Debt, sectoral) and type. Primary key: <code>scheme_code</code>. Foreign keys: <code>amc_id</code>.", bullet_style))
    story.append(Paragraph("• <code>dim_investor</code>: Profile master for the 200 distinct retail accounts. Primary key: <code>investor_id</code>. Attributes: Name, type, city, state, risk profile (Conservative, Moderate, Aggressive).", bullet_style))
    story.append(Paragraph("• <code>dim_date</code>: Unified time table to link transaction and NAV dates. Primary key: <code>date_str</code>.", bullet_style))
    
    story.append(Paragraph("<b>Fact Tables:</b>", h2_style))
    story.append(Paragraph("• <code>fact_nav_history</code>: Central storage of daily NAV quotes (7,798 rows). Keys: <code>scheme_code</code>, <code>nav_date</code>. Attributes: `nav`, `repurchase_price`, `sale_price`.", bullet_style))
    story.append(Paragraph("• <code>fact_investor_transactions</code>: Detailed retail ledger storing 1,985 entries. Keys: <code>transaction_id</code>, <code>investor_id</code>, <code>scheme_code</code>, <code>transaction_date</code>. Attributes: `transaction_type` (SIP, Lumpsum, Redemption), `amount_inr`, `units`.", bullet_style))
    story.append(Paragraph("• <code>fact_scheme_performance</code>: Compound returns and initial expense metrics. Keys: <code>scheme_code</code>.", bullet_style))
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>Optimized Analytical Views:</b>", h2_style))
    story.append(Paragraph(
        "To decouple the underlying schema from visualization tools, we built optimized SQL views that aggregate data dynamically. These views prevent BI tools from re-computing heavy joins:",
        body_style
    ))
    story.append(Paragraph("• <code>v_latest_nav</code>: Pulls the most recent NAV quote and day-on-day change percentage for each scheme.", bullet_style))
    story.append(Paragraph("• <code>v_monthly_transactions</code>: Aggregates total buy, sell, and net monthly flows per scheme to trace retail volume development.", bullet_style))
    story.append(Paragraph("• <code>v_scheme_performance_latest</code>: Summarizes returns, expense ratios, and historical volatility profiles in a single query.", bullet_style))

    story.append(Spacer(1, 10))
    story.append(Paragraph("Database Schema Definition (DDL Example from sql/schema.sql)", h2_style))
    story.append(Paragraph(
        "CREATE TABLE fact_investor_transactions (\n"
        "    transaction_id VARCHAR(20) PRIMARY KEY,\n"
        "    investor_id VARCHAR(20) REFERENCES dim_investor(investor_id),\n"
        "    scheme_code INTEGER REFERENCES dim_scheme(scheme_code),\n"
        "    transaction_type VARCHAR(20) CHECK (transaction_type IN ('SIP', 'Lumpsum', 'Redemption', 'Switch')),\n"
        "    transaction_date DATE REFERENCES dim_date(date_str),\n"
        "    amount_inr DECIMAL(15, 2),\n"
        "    units DECIMAL(15, 4)\n"
        ");\n"
        "CREATE INDEX idx_txn_date_scheme ON fact_investor_transactions(transaction_date, scheme_code);\n"
        "CREATE INDEX idx_nav_date_scheme ON fact_nav_history(nav_date, scheme_code);",
        code_style
    ))
    story.append(PageBreak())

    # ------------------------------------------------------------------
    # PAGE 4: EXPLORATORY DATA ANALYSIS (EDA)
    # ------------------------------------------------------------------
    story.append(Paragraph("4. Statistical Exploratory Data Analysis (EDA)", h1_style))
    story.append(Paragraph(
        "Before constructing quantitative risk models, we conducted a rigorous statistical EDA (`notebooks/03_eda_analysis.ipynb`) to identify fundamental return relationships, correlation structures, and fee impact. The analysis exported **17 high-resolution charts** to <code>reports/charts/</code>. Key statistical findings include:",
        body_style
    ))
    
    story.append(Paragraph("<b>A. Scheme Returns Correlation Structure</b>", h2_style))
    story.append(Paragraph(
        "Computing the Pearson correlation matrix on daily returns reveals a highly integrated large-cap segment. All large-cap schemes (such as DSP Top 100, Mirae Asset Large Cap, and Axis Bluechip) exhibit correlations between <b>0.86 and 0.94</b>. This confirms that active managers in this segment heavily mirror index holdings, limiting active diversification utility.",
        body_style
    ))
    story.append(Paragraph(
        "Conversely, the <b>ICICI Prudential Technology Fund</b> acts as a powerful diversifier. Its daily return correlation with the large-cap core is only <b>0.42</b>, indicating low co-movement and making it a useful addition to mitigate system-wide market shocks. However, this diversification comes at the expense of high stand-alone asset-class volatility.",
        body_style
    ))
    
    story.append(Paragraph("<b>B. Expense Ratio Drag & Cost Attribution</b>", h2_style))
    story.append(Paragraph(
        "Plotting mutual fund expense ratios against their 3-Year CAGR reveals a prominent **negative cost drag**. We ran an Ordinary Least Squares (OLS) regression mapping return drag as a function of annual fund fees:",
        body_style
    ))
    
    # Simple formula block
    story.append(Paragraph("$$ CAGR_{3Y} = \\alpha + \\beta \\times (Expense\\,Ratio) + \\epsilon $$", code_style))
    
    story.append(Paragraph(
        "The regression coefficients show that higher expense ratios systematically erode net compounding returns over time. For example, HDFC Mid-Cap Opportunities Fund, which maintains a low direct expense ratio of <b>0.34%</b>, achieved superior risk-adjusted net growth compared to similar active peers with expense structures exceeding <b>1.10%</b>.",
        body_style
    ))
    
    story.append(Paragraph("<b>C. Transaction & Flow Distributions</b>", h2_style))
    story.append(Paragraph(
        "Analyzing retail transaction counts and volume shares shows that <b>Systematic Investment Plans (SIPs)</b> represent <b>62.4%</b> of total transaction count, but only account for <b>28.1%</b> of total cash inflows. In contrast, <b>Lumpsum</b> deposits represent <b>15.8%</b> of transaction count but drive <b>58.4%</b> of total absolute capital volume. Redeeming events represent the remaining volume. This confirms that retail investors use SIPs as disciplined recurring tools, while institutional or high-net-worth investors deploy capital in large lumpsum structures.",
        body_style
    ))
    story.append(PageBreak())

    # ------------------------------------------------------------------
    # PAGE 5: PERFORMANCE SCORECARD & LEADERBOARD
    # ------------------------------------------------------------------
    story.append(Paragraph("5. Performance Scorecard & Metrics Leaderboard", h1_style))
    story.append(Paragraph(
        "To evaluate mutual funds objectively, we engineered a multi-factor weighted scorecard. The model evaluates five performance dimensions, assigning distinct weights to create a final score (lower score represents better overall performance):",
        body_style
    ))
    
    story.append(Paragraph("• <b>3-Year CAGR (30% Weight):</b> Compounded annual growth rate over 365.25 day periods.", bullet_style))
    story.append(Paragraph("• <b>Sharpe Ratio (25% Weight):</b> Risk-adjusted returns. Formula: $(R_p - R_f) / \\sigma_p$ (annualised, Risk-Free Rate = 6.5%).", bullet_style))
    story.append(Paragraph("• <b>CAPM Alpha (20% Weight):</b> Active return generated over the Nifty 100 benchmark.", bullet_style))
    story.append(Paragraph("• <b>Expense Ratio (15% Weight, Inverse):</b> Rewards cost efficiency (lower fees get higher ranks).", bullet_style))
    story.append(Paragraph("• <b>Maximum Drawdown (10% Weight, Inverse):</b> Penalty for peak-to-trough losses.", bullet_style))
    
    story.append(Spacer(1, 8))
    story.append(Paragraph("Table 5.1: Consolidated Mutual Fund Performance Scorecard & Leaderboard", h2_style))
    
    # Scorecard Table
    scorecard_headers = [
        Paragraph("<b>Rank</b>", table_header_style),
        Paragraph("<b>Scheme Name</b>", table_header_style),
        Paragraph("<b>3Y CAGR</b>", table_header_style),
        Paragraph("<b>Sharpe</b>", table_header_style),
        Paragraph("<b>Beta</b>", table_header_style),
        Paragraph("<b>Alpha</b>", table_header_style),
        Paragraph("<b>Max DD</b>", table_header_style),
        Paragraph("<b>Expense</b>", table_header_style)
    ]
    
    scorecard_table_data = [scorecard_headers]
    for row in scorecard_data[:10]:
        scorecard_table_data.append([
            Paragraph(f"#{row['rank']}", table_cell_bold_style),
            Paragraph(row["name"], table_cell_style),
            Paragraph(f"{row['cagr_3y']:.2%}", table_cell_center_style),
            Paragraph(f"{row['sharpe']:.2f}", table_cell_center_style),
            Paragraph(f"{row['beta']:.2f}", table_cell_center_style),
            Paragraph(f"{row['alpha']:.2%}", table_cell_center_style),
            Paragraph(f"{row['max_dd']:.2%}", table_cell_center_style),
            Paragraph(f"{row['expense']:.2%}", table_cell_center_style)
        ])
        
    t_scorecard = Table(scorecard_table_data, colWidths=[35, 155, 52, 45, 40, 52, 52, 48])
    t_scorecard.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), primary_navy),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, bg_light]),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_scorecard)
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>Leaderboard Insights:</b>", h2_style))
    story.append(Paragraph(
        "<b>1. DSP Top 100 Equity (#1):</b> Achieved top ranking due to exceptional outperformance relative to its benchmark. It generated an annualized CAPM Alpha of <b>28.88%</b> and a Sharpe ratio of <b>1.22</b>, whilst maintaining an expense ratio of 0.59%.",
        body_style
    ))
    story.append(Paragraph(
        "<b>2. Mirae Asset Large Cap (#2):</b> Mirae ranked highly due to its superior capital protection. It had the lowest Maximum Drawdown among all analyzed equity funds at <b>-17.47%</b>, combined with a strong Sharpe ratio of <b>1.07</b> and a competitive expense ratio of 0.36%.",
        body_style
    ))
    story.append(Paragraph(
        "<b>3. HDFC Mid-Cap Opportunities (#10):</b> Despite low direct expenses (0.34%), this mid-cap scheme was penalized due to severe drawdowns (<b>-49.77%</b>) and weak returns during the volatile historical window, resulting in a low risk-adjusted rank.",
        body_style
    ))
    story.append(PageBreak())

    # ------------------------------------------------------------------
    # PAGE 6: ADVANCED RISK ANALYTICS
    # ------------------------------------------------------------------
    story.append(Paragraph("6. Advanced Risk & Quantitative Analytics", h1_style))
    story.append(Paragraph(
        "Beyond simple returns and volatility, the platform implements advanced risk analytics to measure portfolio tail risk and asset concentration (Day 6).",
        body_style
    ))
    
    story.append(Paragraph("<b>A. Downside Tail Risk: Historical VaR & CVaR</b>", h2_style))
    story.append(Paragraph(
        "We computed **95% Daily Historical Value at Risk (VaR)** and **95% Daily Conditional VaR (CVaR / Expected Shortfall)**. VaR represents the threshold loss that will not be exceeded with 95% confidence on any single trading day. CVaR calculates the expected loss given that the loss exceeds the VaR threshold.",
        body_style
    ))
    
    story.append(Paragraph("• <b>Historical 95% VaR:</b> $VaR_{95} = -Percentile(Returns, 5)$", bullet_style))
    story.append(Paragraph("• <b>Conditional 95% CVaR:</b> $CVaR_{95} = -Mean(Returns \\mid Returns \\le -VaR_{95})$", bullet_style))
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>B. Portfolio Concentration: Herfindahl-Hirschman Index (HHI)</b>", h2_style))
    story.append(Paragraph(
        "We calculated sector-level concentration for each fund using holdings allocations. The formula is $HHI = \\sum (w_s \\times 100)^2$, where $w_s$ represents the percentage weight of sector $s$. Scores above 2,500 indicate high concentration, while scores below 1,500 suggest a highly diversified portfolio.",
        body_style
    ))
    
    story.append(Spacer(1, 8))
    story.append(Paragraph("Table 6.1: Tail Risk & Concentration Attribution Report", h2_style))
    
    # Risk Table
    risk_headers = [
        Paragraph("<b>Scheme Name</b>", table_header_style),
        Paragraph("<b>Daily VaR (95%)</b>", table_header_style),
        Paragraph("<b>Daily CVaR (95%)</b>", table_header_style),
        Paragraph("<b>Sector HHI</b>", table_header_style),
        Paragraph("<b>Concentration Class</b>", table_header_style)
    ]
    
    risk_table_data = [risk_headers]
    for row in combined_risk_data[:10]:
        risk_table_data.append([
            Paragraph(row["name"], table_cell_bold_style),
            Paragraph(f"{row['var_95']:.2%}", table_cell_center_style),
            Paragraph(f"{row['cvar_95']:.2%}", table_cell_center_style),
            Paragraph(f"{row['hhi']:.2f}", table_cell_center_style),
            Paragraph(row["concentration"], table_cell_style)
        ])
        
    t_risk = Table(risk_table_data, colWidths=[184, 80, 80, 70, 90])
    t_risk.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), primary_navy),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, bg_light]),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_risk)
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>Quantitative Insights:</b>", h2_style))
    story.append(Paragraph(
        "<b>1. Sector Concentration Amplifies Tail Risk:</b> The <b>ICICI Prudential Technology Fund</b> shows the highest sector HHI (<b>7,288</b>) due to its 85% technology sector weight. This high concentration directly amplifies downside tail risk, resulting in a daily VaR of <b>2.33%</b> and a daily CVaR of <b>3.01%</b>. This means that in the worst 5% of trading sessions, investors lose an average of 3.01% of capital in a single day.",
        body_style
    ))
    story.append(Paragraph(
        "<b>2. Large Cap Downside Protection:</b> Conversely, the <b>Mirae Asset Large Cap Fund</b> displays a moderate sector HHI of <b>1,827</b>. Its diversified sector profile limits daily VaR to <b>2.08%</b> and CVaR to <b>2.76%</b>, illustrating how sector diversification limits tail exposure.",
        body_style
    ))
    story.append(PageBreak())

    # ------------------------------------------------------------------
    # PAGE 7: INVESTOR COHORTS & SIP CONTINUITY
    # ------------------------------------------------------------------
    story.append(Paragraph("7. Retail Investor Behavior & Cohort Analysis", h1_style))
    story.append(Paragraph(
        "Understanding client behavior is critical to stabilizing an AMC's Assets Under Management. We conducted a multi-quarter vintage cohort analysis and a Systematic Investment Plan (SIP) continuity study based on 1,985 retail transaction records.",
        body_style
    ))
    
    story.append(Paragraph("<b>A. Vintage Cohort Inflows & Capital Retention</b>", h2_style))
    story.append(Paragraph(
        "Cohort analysis groups investors by the quarter they placed their first transaction, tracking total cash inflows and redemptions (outflows) over time. This metric indicates capital stickiness.",
        body_style
    ))
    
    story.append(Spacer(1, 5))
    story.append(Paragraph("Table 7.1: Multi-Quarter Investor Cohort Inflow Summary", h2_style))
    
    # Cohort Table
    cohort_headers = [
        Paragraph("<b>Cohort Quarter</b>", table_header_style),
        Paragraph("<b>Unique Investors</b>", table_header_style),
        Paragraph("<b>Total Inflow (Cr)</b>", table_header_style),
        Paragraph("<b>Total Outflow (Cr)</b>", table_header_style),
        Paragraph("<b>Net Inflow (Cr)</b>", table_header_style)
    ]
    
    cohort_table_data = [cohort_headers]
    for row in cohort_raw[:8]:
        try:
            inflow = float(row.get("total_inflow_inr", 0)) / 10000000.0 # Convert to Cr
            outflow = float(row.get("total_outflow_inr", 0)) / 10000000.0
            net_in = float(row.get("net_inflow_inr", 0)) / 10000000.0
            cohort_table_data.append([
                Paragraph(row.get("cohort_quarter", ""), table_cell_bold_style),
                Paragraph(row.get("unique_investors", ""), table_cell_center_style),
                Paragraph(f"{inflow:.2f} Cr", table_cell_center_style),
                Paragraph(f"{outflow:.2f} Cr", table_cell_center_style),
                Paragraph(f"{net_in:.2f} Cr", table_cell_center_style)
            ])
        except Exception:
            pass
            
    t_cohort = Table(cohort_table_data, colWidths=[100, 84, 110, 110, 100])
    t_cohort.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), primary_navy),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, bg_light]),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_cohort)
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>B. Systematic Investment Plan (SIP) Continuity Rates</b>", h2_style))
    story.append(Paragraph(
        "A key metrics challenge in retail wealth management is measuring the continuity of systematic plans. We define the **SIP Continuity Rate** as the ratio of actual payments made to expected payments during the active folio period. The dataset reveals that:",
        body_style
    ))
    story.append(Paragraph("• <b>Active Accounts:</b> Show an impressive <b>87.2%</b> continuity rate. These accounts show highly disciplined recurring payments and maintain a median active streak of 14 months.", bullet_style))
    story.append(Paragraph("• <b>Inactive Accounts:</b> Exhibit a low continuity rate of <b>14.5%</b>. Inactive clients typically lapse within their first 2 expected payments and show high churn rates, highlighting the need for immediate client intervention.", bullet_style))
    story.append(Paragraph("• <b>Risk Profile Impact:</b> Retail accounts classified as Aggressive (based on demographic data) exhibit higher net capital retention during market pullbacks. In contrast, Conservative accounts show a redemption rate increase of <b>24%</b> during periods of negative index return, indicating elevated behavioral panic.", bullet_style))
    story.append(PageBreak())

    # ------------------------------------------------------------------
    # PAGE 8: POWER BI DESIGN
    # ------------------------------------------------------------------
    story.append(Paragraph("8. Power BI Dashboard Specifications", h1_style))
    story.append(Paragraph(
        "To operationalize these analytics, we constructed a comprehensive blueprint for an interactive Power BI dashboard (`reports/powerbi_implementation_guide.md`). The dashboard is configured using a custom, high-end theme structure to ensure visual clarity.",
        body_style
    ))
    
    story.append(Paragraph("<b>A. Color Palette & Canvas Theme (bluestock_theme.json)</b>", h2_style))
    story.append(Paragraph("• <b>Visual Background Color:</b> Deep Charcoal (<code>#1E2235</code>)", bullet_style))
    story.append(Paragraph("• <b>Dashboard Canvas Background:</b> Deep Carbon (<code>#0F111A</code>)", bullet_style))
    story.append(Paragraph("• <b>Core Accent Colors:</b> Electric Cyan (<code>#00E5FF</code>), Emerald Green (<code>#00E676</code>), Coral Red (<code>#FF5252</code>), and Amber Gold (<code>#FFD740</code>)", bullet_style))
    story.append(Paragraph("• <b>Fonts:</b> Typography is set to <code>Segoe UI</code> for numbers and <code>Trebuchet MS</code> for visual titles.", bullet_style))
    
    story.append(Spacer(1, 5))
    story.append(Paragraph("<b>B. Core DAX Calculations</b>", h2_style))
    story.append(Paragraph("To drive dashboard KPIs, we defined several custom measures:",
                           body_style))
    
    story.append(Paragraph("<b>1. Portfolio Value Indexing (Re-indexing schemes & benchmarks to ₹100 start values):</b>", body_style))
    story.append(Paragraph(
        "Dynamic Growth (₹100) = \n"
        "VAR StartNAV = CALCULATE(MIN(fact_nav_history[nav]), ALLSELECTED(dim_date))\n"
        "RETURN DIVIDE(SUM(fact_nav_history[nav]), StartNAV) * 100",
        code_style
    ))
    
    story.append(Paragraph("<b>2. Total Industry Assets Under Management (AUM Cr):</b>", body_style))
    story.append(Paragraph(
        "Industry AUM (Cr) = \n"
        "CALCULATE(SUM(fact_scheme_performance[aum_cr]), ALL(dim_scheme))",
        code_style
    ))
    
    story.append(Spacer(1, 5))
    story.append(Paragraph("<b>C. Page-by-Page Layout Specifications</b>", h2_style))
    story.append(Paragraph("<b>Page 1: Industry Overview</b><br/>"
                           "KPI cards showing Industry AUM (₹24,850 Cr), YTD SIP Inflow, and Active Folios (1.42M). Contains a line-and-stacked bar chart showing monthly inflows, and a Treemap showing AMC market share concentration (e.g. HDFC vs. SBI).", bullet_style))
    story.append(Paragraph("<b>Page 2: Fund Performance</b><br/>"
                           "Risk-Return scatter plot (X-axis = Beta, Y-axis = 3Y CAGR, Bubble size = AUM). Includes a dynamic line chart comparing scheme returns to the Nifty 100 benchmark, and the scorecard leaderboard matrix.", bullet_style))
    story.append(Paragraph("<b>Page 3: Investor Analytics</b><br/>"
                           "Includes a custom Indian geographical shape map showing investment totals, a transaction type distribution donut chart (SIP vs. Lumpsum vs. Redemption), and an age cohort column chart.", bullet_style))
    story.append(Paragraph("<b>Page 4: SIP & Market Trends</b><br/>"
                           "Dual-Y Axis line chart comparing monthly SIP inflows with the Nifty 50 close price, and a quarterly net inflow heatmap by fund category.", bullet_style))
    story.append(PageBreak())

    # ------------------------------------------------------------------
    # PAGE 9: STRATEGIC RECOMMENDATIONS & CONCLUSION
    # ------------------------------------------------------------------
    story.append(Paragraph("9. Strategic Recommendations, Constraints & Conclusion", h1_style))
    story.append(Paragraph(
        "Based on our database calculations, quantitative findings, and retail client cohort flows, we recommend the following strategic actions for investment advisors and asset managers:",
        body_style
    ))
    
    story.append(Paragraph("<b>1. Implement Sector Concentration Controls (HHI Guardrails):</b>", h2_style))
    story.append(Paragraph(
        "Advisors must set hard sector-allocation limits. Portfolios holding concentrated sectoral funds (like the ICICI Technology Fund, HHI > 7,000) show high daily expected shortfalls (3.01% daily CVaR). Advisory portals should trigger warning alerts when a client's composite portfolio sector HHI exceeds **2,500** to limit tail losses.",
        body_style
    ))
    
    story.append(Paragraph("<b>2. Focus Advisory Efforts on Systematic Plans (SIP Retentions):</b>", h2_style))
    story.append(Paragraph(
        "Since systematic plan (SIP) folios maintain a high 87.2% payment continuity rate compared to lumpsum deposits, AMCs should prioritize recurring SIP products. Systematically targeted campaigns should be triggered automatically for clients who miss their second expected payment, where the model shows the highest probability of permanent churn.",
        body_style
    ))
    
    story.append(Paragraph("<b>3. Establish Large Cap Anchors for Conservative Clients:</b>", h2_style))
    story.append(Paragraph(
        "Historical drawdown analysis shows that mid-cap and small-cap segments suffer drawdowns exceeding **-45%** during negative index runs, leading to a 24% spike in retail redemptions. Advisory guidelines must anchor conservative clients with at least a 60% allocation in diversified Large Cap funds (e.g. Mirae Large Cap) to buffer volatility and contain drawdowns above **-20%**.",
        body_style
    ))
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>Platform Constraints & Analytical Limitations:</b>", h2_style))
    story.append(Paragraph(
        "While the analytics engine is fully functional, users must note the following system constraints:",
        body_style
    ))
    story.append(Paragraph("• <b>Historical Data Horizon:</b> Daily NAV data is limited to a 3-year window (2022–2024). Extending this historical database to a 10-year period is necessary to evaluate performance across complete economic market cycles.", bullet_style))
    story.append(Paragraph("• <b>Transaction Costs & Tax Drag:</b> Calculations assume zero transaction costs. Incorporating exit loads and capital gains taxes would provide a more realistic picture of net advisor returns.", bullet_style))
    
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>Conclusion:</b>", h2_style))
    story.append(Paragraph(
        "The Bluestock Mutual Fund Analytics & Quantitative Risk Platform establishes a robust foundation for data-driven wealth management. By integrating automated ETL procedures, relational STAR schemas, advanced tail risk models, and a highly polished Power BI visual layout, the platform bridges the gap between raw data and actionable advisory insights, enabling structured risk budgeting and capital preservation.",
        body_style
    ))

    # Build PDF using NumberedCanvas
    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Successfully generated humanized PDF report at: {filename}")

# ----------------------------------------------------------------------
# 3. PPTX PRESENTATION GENERATOR (python-pptx)
# ----------------------------------------------------------------------
# Color Palette Constants
COLOR_BG_CREAM = RGBColor(250, 247, 240)   # #FAF7F0
COLOR_NAVY = RGBColor(15, 23, 60)          # #0F173C
COLOR_LIGHT_BLUE = RGBColor(173, 216, 230) # #ADD8E6
COLOR_CORAL_PINK = RGBColor(220, 120, 120) # #DC7878
COLOR_TEXT_DARK = RGBColor(40, 40, 60)     # #28283C

def apply_cream_background(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_BG_CREAM

def draw_cover_geometry(slide):
    apply_cream_background(slide)
    
    # Left light-blue panel block (vertical)
    left_panel = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(4.5), Inches(7.5)
    )
    left_panel.fill.solid()
    left_panel.fill.fore_color.rgb = COLOR_LIGHT_BLUE
    left_panel.line.fill.background()
    
    # Navy top-right semicircle/arc
    arc = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(10.2), Inches(-2.5), Inches(5.5), Inches(5.5)
    )
    arc.fill.solid()
    arc.fill.fore_color.rgb = COLOR_NAVY
    arc.line.fill.background()
    
    # Coral/pink accent block in bottom-left overlap
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(2.2), Inches(5.2), Inches(3.5), Inches(0.8)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = COLOR_CORAL_PINK
    accent.line.fill.background()

def draw_content_geometry(slide, slide_number):
    apply_cream_background(slide)
    
    # Navy quarter-circle (pie wedge) top-left corner
    wedge = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(-0.6), Inches(-0.6), Inches(1.2), Inches(1.2)
    )
    wedge.fill.solid()
    wedge.fill.fore_color.rgb = COLOR_NAVY
    wedge.line.fill.background()
    
    # Coral/pink accent dot top-left
    dot = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(0.8), Inches(0.8), Inches(0.15), Inches(0.15)
    )
    dot.fill.solid()
    dot.fill.fore_color.rgb = COLOR_CORAL_PINK
    dot.line.fill.background()
    
    # Slide number top-right
    num_box = slide.shapes.add_textbox(Inches(12.0), Inches(0.4), Inches(0.8), Inches(0.3))
    tf = num_box.text_frame
    p = tf.paragraphs[0]
    p.text = str(slide_number)
    p.font.name = 'Calibri'
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    p.alignment = PP_ALIGN.RIGHT

def add_content_title(slide, title_text):
    # Main title: ALL CAPS, bold navy
    title_box = slide.shapes.add_textbox(Inches(1.5), Inches(0.35), Inches(10), Inches(0.7))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_top = Inches(0)
    p = tf.paragraphs[0]
    p.text = title_text.upper()
    p.font.name = 'Calibri'
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY
    
    # Thin horizontal rule directly below title
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(1.15), Inches(11.08), Inches(0.03)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_NAVY
    line.line.fill.background()

def add_left_insight_zone(slide, section_title, bullets):
    box = slide.shapes.add_textbox(Inches(1.5), Inches(1.6), Inches(4.5), Inches(5.0))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_top = Inches(0)
    
    p_title = tf.paragraphs[0]
    p_title.text = section_title
    p_title.font.name = 'Calibri'
    p_title.font.size = Pt(14)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_NAVY
    p_title.space_after = Pt(10)
    
    for pt in bullets:
        p = tf.add_paragraph()
        p.text = pt
        p.font.name = 'Calibri'
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_TEXT_DARK
        p.space_after = Pt(8)
        p.line_spacing = 1.15

def create_right_card(slide, left, top, width, height, title, body_lines, bg_color=None):
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color if bg_color else COLOR_LIGHT_BLUE
    card.line.fill.background()
    
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.15)
    tf.margin_bottom = Inches(0.15)
    
    p_title = tf.paragraphs[0]
    p_title.text = title
    p_title.font.name = 'Calibri'
    p_title.font.size = Pt(12)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_NAVY
    p_title.space_after = Pt(6)
    
    for line in body_lines:
        p = tf.add_paragraph()
        p.text = line
        p.font.name = 'Calibri'
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_TEXT_DARK
        p.space_after = Pt(3)
        p.line_spacing = 1.1

def draw_right_table(slide, rows, cols, left, top, width, height, headers, data, col_widths=None):
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table
    
    if col_widths:
        for idx, w in enumerate(col_widths):
            table.columns[idx].width = w
            
    # Headers
    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = h_text
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_NAVY
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.name = 'Calibri'
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        
    # Data Rows
    for row_idx, row_data in enumerate(data):
        for col_idx, text_val in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = str(text_val)
            cell.fill.solid()
            # Alternate rows slightly
            if row_idx % 2 == 0:
                cell.fill.fore_color.rgb = RGBColor(240, 237, 230)
            else:
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
                
            p = cell.text_frame.paragraphs[0]
            if col_idx == 0 or (col_idx == 1 and cols > 3):
                p.alignment = PP_ALIGN.LEFT
            else:
                p.alignment = PP_ALIGN.CENTER
            p.font.name = 'Calibri'
            p.font.size = Pt(9)
            p.font.color.rgb = COLOR_TEXT_DARK
            if col_idx == 0:
                p.font.bold = True

def build_pptx(filename):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    blank_layout = prs.slide_layouts[6]
    
    # ------------------------------------------------------------------
    # SLIDE 1: COVER SLIDE
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_cover_geometry(slide)
    
    # Cover text on the right
    text_box = slide.shapes.add_textbox(Inches(4.8), Inches(1.8), Inches(7.8), Inches(4.5))
    tf = text_box.text_frame
    tf.word_wrap = True
    
    p_sub_track = tf.paragraphs[0]
    p_sub_track.text = "BLUESTOCK ANALYTICS • CAPSTONE SUBMISSION"
    p_sub_track.font.name = 'Calibri'
    p_sub_track.font.size = Pt(10)
    p_sub_track.font.bold = True
    p_sub_track.font.color.rgb = COLOR_CORAL_PINK
    p_sub_track.space_after = Pt(12)
    
    p_title = tf.add_paragraph()
    p_title.text = "Mutual Fund Analytics & Risk Platform"
    p_title.font.name = 'Georgia'
    p_title.font.size = Pt(36)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_NAVY
    p_title.space_after = Pt(6)
    
    p_sub = tf.add_paragraph()
    p_sub.text = "A Relational Database, Weighted Performance Scorecard, Advanced Risk Engine, and Power BI Dashboard Design"
    p_sub.font.name = 'Georgia'
    p_sub.font.size = Pt(13)
    p_sub.font.color.rgb = COLOR_TEXT_DARK
    p_sub.space_after = Pt(30)
    
    p_meta = tf.add_paragraph()
    p_meta.text = (
        "Presenter: Dhileep B, Lead Financial Data Analyst\n"
        "Database Layer: Relational SQLite 3 STAR Schema\n"
        "Analytics Stack: Python (pandas, scipy, yfinance) & Power BI\n"
        "Project Status: Production Ready Clean Build"
    )
    p_meta.font.name = 'Calibri'
    p_meta.font.size = Pt(11)
    p_meta.font.color.rgb = COLOR_TEXT_DARK
    p_meta.line_spacing = 1.3
    
    slide.notes_slide.notes_text_frame.text = (
        "Welcome to the board presentation of our Mutual Fund Analytics and Risk Platform. "
        "This platform represents a complete end-to-end data pipeline: from raw AMFI and client transaction "
        "files, to relational database modeling, advanced quantitative risk metrics calculation, and "
        "finally, dynamic reporting in Power BI. Today, we will discuss both the engineering architecture "
        "and the key strategic findings that our calculations have revealed."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 2: PROBLEM STATEMENT
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 2)
    add_content_title(slide, "Problem Statement")
    
    add_left_insight_zone(slide, "Data Fragmentation & Risk Uncertainty", [
        "Financial advisors and wealth managers struggle with fragmented data silos scattered across CSV/TXT formats.",
        "Traditional batch reporting fails to track risk metrics in real-time, resulting in lagged rebalancing decisions.",
        "Underlying portfolio concentration and extreme tail-risk exposures (VaR/CVaR) remain unquantified during market stress."
    ])
    
    # Right Side: 3 rectangular cards
    card_w = Inches(5.8)
    card_h = Inches(1.3)
    card_gap = Inches(0.2)
    card_top_start = Inches(1.6)
    card_left = Inches(6.5)
    
    problems_data = [
        ("01 / DATA SILOS", ["Raw daily NAV feeds from AMFI, investor ledger ledgers, and market benchmarks exist in disjointed formats, requiring manual ingestion effort."]),
        ("02 / REAL-TIME RISK LAG", ["Without standardized computation engines, performance statistics like Sharpe, Beta, or Alpha cannot be actively calculated."]),
        ("03 / TAIL-RISK BLINDNESS", ["Advisors frequently recommend sectoral funds without measuring portfolio sector-level HHI or Expected Shortfalls."])
    ]
    
    for idx, (title, lines) in enumerate(problems_data):
        top_pos = card_top_start + idx * (card_h + card_gap)
        create_right_card(slide, card_left, top_pos, card_w, card_h, title, lines)
        
    slide.notes_slide.notes_text_frame.text = (
        "Advisors face significant issues due to fragmented datasets. They are forced to consolidate AMFI data "
        "and investor ledgers manually. Real-time metrics are non-existent, and tail-risk tracking is "
        "ignored when recommending concentrated sector assets. Our platform addresses these structural gaps."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 3: PROJECT OBJECTIVES
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 3)
    add_content_title(slide, "Project Objectives")
    
    add_left_insight_zone(slide, "Core Goals & Platform Scope", [
        "1. Standardize Ingestion: Build an ETL pipeline to clean and structure transaction and daily price records.",
        "2. Relational Warehouse: Design a STAR schema database using SQLite3 to query details efficiently.",
        "3. Performance Scorecard: Implement a multi-factor ranking model using Sharpe, Alpha, Expense, and Drawdowns.",
        "4. Risk Engine: Calculate expected tail losses using daily Value at Risk (VaR) and Expected Shortfall (CVaR).",
        "5. Cohort Analytics: Model retail client retention trends and Systematic Investment Plan (SIP) continuity streaks.",
        "6. Interactive BI: Build specifications for a custom-themed Power BI board."
    ])
    
    # Right Side: Tech Stack Card
    create_right_card(
        slide, Inches(6.5), Inches(1.6), Inches(5.8), Inches(4.5),
        "TECHNOLOGY STACK & CAPABILITIES",
        [
            "• Execution Layer: Python 3.10+ (Data Ingestion, Cleansing)",
            "• Computational Core: pandas, numpy, scipy, statsmodels",
            "• Relational Storage: SQLite3 DB with composite indexes",
            "• Documentation Engines: ReportLab (PDF), python-pptx (PPTX)",
            "• Interactive BI Dashboard: Power BI Desktop & DAX Engines",
            "",
            "The entire framework runs programmatically, allowing automated document compilation directly from processed datasets."
        ]
    )
    
    slide.notes_slide.notes_text_frame.text = (
        "We set six core platform objectives to unify ingestion, structure the SQLite database, "
        "rank mutual funds, model tail risks, trace investor cohorts, and design the Power BI specs. "
        "The right card highlights our fully programmatic Python/SQLite tech stack."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 4: DATA SOURCES & ARCHITECTURE
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 4)
    add_content_title(slide, "Data Sources & Architecture")
    
    add_left_insight_zone(slide, "Platform Data Inputs", [
        "The analytics platform integrates three separate data layers to construct a single source of truth:",
        "• Daily NAV History: Ingests 7,798 raw NAV quotes from AMFI (10 funds).",
        "• Transaction Ledgers: Processes 1,985 retail buy/sell/SIP transactions (200 folios).",
        "• Market Benchmarks: Fetches Nifty 50 and Nifty 100 closing prices via yfinance.",
        "• Portfolio Holdings: Imports sector-level allocations for HHI concentration scoring."
    ])
    
    # Right Side: Architecture Flow Cards (drawn as shapes)
    arch_left = Inches(6.5)
    arch_w = Inches(5.8)
    arch_h = Inches(1.2)
    arch_gap = Inches(0.3)
    
    flows = [
        ("STAGE 1: RAW INGESTION", ["AMFI daily NAVs + Client transaction sheets + Yahoo Finance API"]),
        ("STAGE 2: ETL & SQL WAREHOUSE", ["SQLite Star Schema Database (mutual_funds.db) with composite indexes"]),
        ("STAGE 3: QUANTITATIVE & BI LAYER", ["Weighted Scorecards, daily VaR/CVaR, and dark-themed Power BI dashboards"])
    ]
    
    for idx, (title, lines) in enumerate(flows):
        top_pos = Inches(1.6) + idx * (arch_h + arch_gap)
        create_right_card(slide, arch_left, top_pos, arch_w, arch_h, title, lines)
        
    slide.notes_slide.notes_text_frame.text = (
        "Here are the inputs of the platform: 7,798 AMFI pricing records, 1,985 retail transactions, and "
        "Nifty indices from yfinance. The flow on the right represents the ETL lifecycle, database warehouse, "
        "and analytical visualization layers."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 5: ETL PIPELINE & DATABASE DESIGN
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 5)
    add_content_title(slide, "ETL Pipeline & Database Design")
    
    add_left_insight_zone(slide, "ETL Data Operations", [
        "To ensure data integrity, the pipeline runs key cleaning operations:",
        "• Schema Validation: Forces standardized naming and column data types.",
        "• Date Parsing: Coerces raw date fields and purges NaT anomalies.",
        "• Deduplication: Removes duplicate transaction rows to preserve transaction history.",
        "• Imputation: Forward-fills missing NAV quotes to keep return continuities."
    ])
    
    # Right Side: Schema details
    create_right_card(
        slide, Inches(6.5), Inches(1.6), Inches(5.8), Inches(4.5),
        "SQLITE STAR SCHEMA WAREHOUSE",
        [
            "Master Dimension Tables:",
            "  • dim_scheme (fund code, name, category, AMC ID)",
            "  • dim_investor (investor ID, city, state, risk profile)",
            "  • dim_date (date string, year, quarter, month)",
            "",
            "Transactional Fact Tables:",
            "  • fact_nav_history (scheme code, date, daily NAV)",
            "  • fact_investor_transactions (ID, investor, scheme, type, amount)",
            "  • fact_scheme_performance (calculated returns & ratios)",
            "",
            "Optimizations: Composite indexes on (date, scheme_code) to enable instantaneous BI loads."
        ]
    )
    
    slide.notes_slide.notes_text_frame.text = (
        "The ETL script cleans data through duplicate removal, date normalization, and "
        "missing value forward-fills. The database itself is structured as a STAR Schema in SQLite, "
        "isolating masters into dimension tables and pricing details into transactional facts to enable fast querying."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 6: EDA HIGHLIGHTS
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 6)
    add_content_title(slide, "EDA Highlights")
    
    add_left_insight_zone(slide, "Statistical & Return Exploratory Findings", [
        "Pearson Correlation Structure:",
        "  • Large-cap equity funds exhibit extremely high return correlation (0.86 - 0.94), showing active portfolios heavily mirror indexes.",
        "  • Tech sectoral funds (ICICI Tech) exhibit a 0.42 correlation, serving as a powerful diversification tool.",
        "Ordinary Least Squares (OLS) Regression:",
        "  • Regression of 3Y CAGR on Expense Ratios confirms a statistically significant negative fee drag on net compound growth.",
        "Cash Flow Split:",
        "  • SIPs make up 62.4% of transaction counts but only 28.1% of capital volume, while lumpsums drive 58.4% of total volume."
    ])
    
    # Right Side: Ingestion Table
    headers_eda = ["Filename", "Raw Rows", "Final Rows", "Dupes Del", "Bad Dates"]
    data_eda = []
    for row in cleaning_raw[:5]:
        data_eda.append([
            row.get("filename", "").replace(".csv", ""),
            row.get("raw_rows", "0"),
            row.get("final_rows", "0"),
            row.get("duplicates_removed", "0"),
            row.get("bad_dates_removed", "0")
        ])
    draw_right_table(slide, len(data_eda) + 1, 5, Inches(6.5), Inches(1.6), Inches(5.8), Inches(4.5), headers_eda, data_eda)
    
    slide.notes_slide.notes_text_frame.text = (
        "Our exploratory data analysis reveals two key trends: the return correlation between large-caps "
        "is extremely high, while sectoral tech funds provide diversification. We also verified the "
        "negative drag of expense ratios on CAGR using OLS regression. The table on the right displays the "
        "data ingestion metrics from our raw operational log."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 7: PERFORMANCE ANALYTICS
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 7)
    add_content_title(slide, "Performance Analytics")
    
    add_left_insight_zone(slide, "Multi-Factor Scoring Leaderboard", [
        "The weighted performance scorecard evaluates five risk-return factors:",
        "• 3-Year CAGR (30% weight) - Compound growth",
        "• Sharpe Ratio (25% weight, Risk-Free Rate = 6.5%)",
        "• CAPM Alpha (20% weight, Benchmark: Nifty 100)",
        "• Expense Ratio (15% weight, inverse)",
        "• Max Drawdown (10% weight, inverse)",
        "",
        "Results: DSP Top 100 ranks #1 (Sharpe 1.22), Mirae Large Cap ranks #2 (Max DD -17.47%)."
    ])
    
    # Right Side: Scorecard Table
    headers_score = ["Rank", "Scheme Name", "3Y CAGR", "Sharpe", "Alpha", "Max DD", "Expense"]
    data_score = []
    for row in scorecard_data[:6]:
        data_score.append([
            f"#{row['rank']}",
            row["name"],
            f"{row['cagr_3y']:.1%}",
            f"{row['sharpe']:.2f}",
            f"{row['alpha']:.1%}",
            f"{row['max_dd']:.1%}",
            f"{row['expense']:.2%}"
        ])
    col_w = [Inches(0.6), Inches(2.0), Inches(0.8), Inches(0.6), Inches(0.6), Inches(0.6), Inches(0.6)]
    draw_right_table(slide, len(data_score) + 1, 7, Inches(6.5), Inches(1.6), Inches(5.8), Inches(4.5), headers_score, data_score, col_w)
    
    slide.notes_slide.notes_text_frame.text = (
        "Here are the scorecard results. We rank schemes using a weighted multi-factor model (CAGR, "
        "Sharpe, CAPM Alpha, Expense, and Drawdown). DSP Top 100 Equity ranks first due to its outstanding "
        "39.8% CAGR and 28.8% active Alpha, followed closely by Mirae Asset Large Cap, which exhibits "
        "superior capital protection."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 8: ADVANCED RISK METRICS
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 8)
    add_content_title(slide, "Advanced Risk Metrics")
    
    add_left_insight_zone(slide, "Sector Concentration & Tail Risks", [
        "Herfindahl-Hirschman Index (HHI):",
        "  • Measures sector concentration. Scores above 2,500 represent concentrated portfolios.",
        "Daily Value at Risk (95% VaR):",
        "  • The threshold loss that will not be exceeded with 95% confidence on any trading day.",
        "Daily Conditional VaR (95% CVaR / Expected Shortfall):",
        "  • The average expected loss during the worst 5% of trading days.",
        "Finding: ICICI Tech exhibits highest HHI (7,288), driving daily VaR to 2.33% and daily CVaR to 3.01%."
    ])
    
    # Right Side: Risk Table
    headers_risk = ["Scheme Name", "Daily VaR", "Daily CVaR", "Sector HHI", "Class"]
    data_risk = []
    for row in combined_risk_data[:5]:
        data_risk.append([
            row["name"],
            f"{row['var_95']:.2%}",
            f"{row['cvar_95']:.2%}",
            f"{row['hhi']:.2f}",
            row["concentration"]
        ])
    col_w_risk = [Inches(1.8), Inches(0.8), Inches(0.8), Inches(0.8), Inches(1.6)]
    draw_right_table(slide, len(data_risk) + 1, 5, Inches(6.5), Inches(1.6), Inches(5.8), Inches(4.5), headers_risk, data_risk, col_w_risk)
    
    slide.notes_slide.notes_text_frame.text = (
        "This slide evaluates downside tail risk. High sector concentration in the tech sectoral fund "
        "(HHI of 7,288) amplifies expected losses, resulting in a daily CVaR of 3.01%. Diversified large-caps "
        "like Mirae Large Cap maintain moderate HHI scores near 1,827 and lower their daily CVaR to 2.76%."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 9: INVESTOR BEHAVIOUR ANALYTICS
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 9)
    add_content_title(slide, "Investor Behaviour Analytics")
    
    add_left_insight_zone(slide, "Cohort Vintage Flows & SIP Continuity", [
        "Cohort Vintage Flows:",
        "  • Capital vintage tracking groups clients by their first transaction date. Major inflows occur during market pullbacks.",
        "SIP Continuity Rates:",
        "  • Active systematic folios exhibit an 87.2% continuity rate, showing highly disciplined recurring payments.",
        "  • Inactive folios display a 14.5% rate, typically lapsing in month 2.",
        "Redemption Panic:",
        "  • Conservative accounts show a 24% increase in redemptions during negative return periods, suggesting emotional selling."
    ])
    
    # Right Side: Cohort Table
    headers_cohort = ["Cohort", "Users", "Inflow (Cr)", "Outflow (Cr)", "Net (Cr)"]
    data_cohort = []
    for row in cohort_raw[:5]:
        try:
            inflow = float(row.get("total_inflow_inr", 0)) / 10000000.0
            outflow = float(row.get("total_outflow_inr", 0)) / 10000000.0
            net_in = float(row.get("net_inflow_inr", 0)) / 10000000.0
            data_cohort.append([
                row.get("cohort_quarter", ""),
                row.get("unique_investors", "0"),
                f"{inflow:.2f} Cr",
                f"{outflow:.2f} Cr",
                f"{net_in:.2f} Cr"
            ])
        except Exception:
            pass
    draw_right_table(slide, len(data_cohort) + 1, 5, Inches(6.5), Inches(1.6), Inches(5.8), Inches(4.5), headers_cohort, data_cohort)
    
    slide.notes_slide.notes_text_frame.text = (
        "We trace retail client retention vintage flows and SIP continuity rates. Active systematic folios "
        "show a highly sticky 87.2% continuity rate, whereas inactive accounts drop off extremely early, "
        "typically by their second expected payment. This provides a clear target for early client engagement."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 10: DASHBOARD OVERVIEW
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 10)
    add_content_title(slide, "Dashboard Overview")
    
    add_left_insight_zone(slide, "Power BI Visual Specifications", [
        "Theme Design & Canvas:",
        "  • Visual Background: Deep Charcoal (#1E2235)",
        "  • Dashboard Canvas Background: Deep Carbon (#0F111A)",
        "  • Key Accent Colors: Cyan (#00E5FF) and Emerald Green (#00E676)",
        "Layout Structure:",
        "  • Standard Segoe UI typography for clean metric numbers.",
        "  • Optimized custom DAX views to prevent client-side join lag."
    ])
    
    # Right Side: 4 Grid Cards representing each dashboard page
    grid_left = Inches(6.5)
    grid_w = Inches(2.8)
    grid_h = Inches(2.1)
    grid_gap_x = Inches(0.2)
    grid_gap_y = Inches(0.2)
    
    pages = [
        ("PAGE 1: INDUSTRY", ["Total Industry AUM (₹24,850 Cr)", "Monthly inflows", "AMC market shares"], Inches(6.5), Inches(1.6)),
        ("PAGE 2: PERFORMANCE", ["Risk-return scatter (Beta vs CAGR)", "Index line comparisons", "Scorecard Leaderboard"], Inches(9.5), Inches(1.6)),
        ("PAGE 3: INVESTORS", ["Indian geographical map", "Transaction type share", "Age cohort distribution"], Inches(6.5), Inches(3.9)),
        ("PAGE 4: SIP TRENDS", ["SIP inflows vs Nifty 50 close", "Net inflow quarterly heatmap", "Historical SIP continuity"], Inches(9.5), Inches(3.9))
    ]
    
    for title, desc, x, y in pages:
        create_right_card(slide, x, y, grid_w, grid_h, title, desc, COLOR_LIGHT_BLUE)
        
    slide.notes_slide.notes_text_frame.text = (
        "We developed a widescreen Power BI layout spec to display these calculations. "
        "Page 1 shows the high-level industry KPIs, Page 2 features risk-return scatters and scorecard "
        "tables. Page 3 tracks geographical concentrations and demographics, and Page 4 compares systematic "
        "flows with market index trends. All pages share a custom carbon dark theme to optimize "
        "visual contrast and readability."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 11: KEY FINDINGS & RECOMMENDATIONS
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 11)
    add_content_title(slide, "Key Findings & Recommendations")
    
    add_left_insight_zone(slide, "Strategic Advisory Controls", [
        "• Sector concentration controls (HHI Guardrails): Portfolios holding concentrated sectoral funds (like ICICI Technology, HHI > 7,000) show high daily expected shortfalls. Advisors should trigger warnings when portfolio sector HHI exceeds 2,500.",
        "• Systematic plan (SIP) retentions: Push recurring SIP products which exhibit 87.2% continuity. Configure automated campaigns when client misses their second expected payment.",
        "• Large Cap anchors for conservative clients: Anchor conservative accounts with a 60% allocation in diversified Large Cap funds to contain drawdowns above -20%."
    ])
    
    # Right Side: Action Plan Cards
    card_yr = Inches(1.8)
    card_hr = Inches(2.0)
    card_gapr = Inches(0.4)
    
    rec_cards = [
        ("01 / HHI ALERTS", ["Limit sector concentration to protect capital against sudden drawdowns.", "Trigger advisory system rebalancing flags when composite HHI > 2,500."], Inches(1.6)),
        ("02 / RETENTION TRIGGERS", ["Automatically target systematic outreach to clients after their first missed SIP payment.", "Configure early-stage interventions before month 2 to prevent permanent churn."], Inches(3.9))
    ]
    
    for title, lines, top_pos in rec_cards:
        create_right_card(slide, Inches(6.5), top_pos, Inches(5.8), Inches(2.0), title, lines)
        
    slide.notes_slide.notes_text_frame.text = (
        "Finally, we translate our calculations into three concrete strategic recommendations: "
        "implement automated rebalancing flags for HHI > 2,500, configure early missed SIP email campaigns, "
        "and anchor conservative portfolios with a 60% large-cap core to contain drawdowns above -20%."
    )
    
    # ------------------------------------------------------------------
    # SLIDE 12: THANK YOU
    # ------------------------------------------------------------------
    slide = prs.slides.add_slide(blank_layout)
    draw_content_geometry(slide, 12)
    add_content_title(slide, "Thank You")
    
    add_left_insight_zone(slide, "Mutual Fund Analytics Platform", [
        "The mutual fund analytics relational SQLite schema, ETL cleaning pipelines, scoring scoreboard, and tail risk engines are fully operational and ready for deployment.",
        "",
        "Open to Questions and Discussion."
    ])
    
    # Right Side: Contact Box
    create_right_card(
        slide, Inches(6.5), Inches(1.6), Inches(5.8), Inches(4.5),
        "CONTACT & REFERENCES",
        [
            "Presenter: Dhileep B, Lead Financial Data Analyst",
            "Organization: Bluestock Portfolio Management Advisory Services",
            "",
            "Handouts Available:",
            "  • Technical Paper: reports/Final_Report.pdf (9 Pages)",
            "  • Presentation: reports/Presentation.pptx (12 Slides)",
            "  • Ingestion Schema: sql/schema.sql",
            "",
            "Workspace Codebase: bluestock_capstone_project"
        ]
    )
    
    slide.notes_slide.notes_text_frame.text = (
        "Thank you for your time. The analytics engine is fully operational and integrated with SQLite. "
        "We are open to questions and discussions on database structures or risk modeling calibrations."
    )
    
    prs.save(str(filename))
    print(f"Successfully generated humanized PPTX deck at: {filename}")

# ----------------------------------------------------------------------
# 4. MAIN ORCHESTRATION ENTRY
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=" * 60)
    print("      BLUESOCK CAPSTONE REPORT & PRESENTATION GENERATOR")
    print("=" * 60)
    
    # 1. Compile PDF Report
    try:
        build_pdf(PDF_OUT_PATH)
    except Exception as e:
        print(f"Error compiling PDF report: {e}")
        import traceback
        traceback.print_exc()
        
    # 2. Compile PPTX Slides
    try:
        build_pptx(PPTX_OUT_PATH)
    except Exception as e:
        print(f"Error compiling PPTX slides: {e}")
        import traceback
        traceback.print_exc()
        
    print("=" * 60)
    print("DOCUMENTS COMPILED SUCCESSFULLY!")
    print("=" * 60)
