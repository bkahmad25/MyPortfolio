const fs = require('fs');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ExternalHyperlink
} = require('docx');

function createATSResumeDocx() {
  const primaryBlue = "1E3A8A";
  const darkText = "0F172A";
  const grayText = "475569";
  const lightBorder = "CBD5E1";

  // Section heading builder
  function createSectionHeading(title) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 200, after: 80 },
      border: {
        bottom: {
          color: primaryBlue,
          space: 2,
          style: BorderStyle.SINGLE,
          size: 12,
        },
      },
      children: [
        new TextRun({
          text: title.toUpperCase(),
          bold: true,
          font: "Calibri",
          size: 22, // 11pt
          color: primaryBlue,
        }),
      ],
    });
  }

  // Bullet point builder
  function createBullet(leadText, normalText) {
    return new Paragraph({
      bullet: { level: 0 },
      spacing: { before: 40, after: 40 },
      children: [
        new TextRun({
          text: leadText,
          bold: true,
          font: "Calibri",
          size: 20, // 10pt
          color: darkText,
        }),
        new TextRun({
          text: normalText,
          font: "Calibri",
          size: 20, // 10pt
          color: "1A1A1A",
        }),
      ],
    });
  }

  // Job / Project header builder
  function createItemHeader(title, company, dateLoc) {
    return new Paragraph({
      spacing: { before: 120, after: 40 },
      children: [
        new TextRun({
          text: title + " | ",
          bold: true,
          font: "Calibri",
          size: 21, // 10.5pt
          color: darkText,
        }),
        new TextRun({
          text: company,
          bold: true,
          font: "Calibri",
          size: 21,
          color: primaryBlue,
        }),
        new TextRun({
          text: "    (" + dateLoc + ")",
          italics: true,
          font: "Calibri",
          size: 19, // 9.5pt
          color: grayText,
        }),
      ],
    });
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Calibri",
            size: 20,
            color: "1A1A1A",
          },
          paragraph: {
            spacing: { line: 276 }, // 1.15 line height
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 720,    // 0.5 in
              right: 720,
              bottom: 720,
              left: 720,
            },
          },
        },
        children: [
          // NAME
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 40 },
            children: [
              new TextRun({
                text: "BILAL AHMED",
                bold: true,
                font: "Calibri",
                size: 36, // 18pt
                color: darkText,
              }),
            ],
          }),

          // SUBTITLE
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 80 },
            children: [
              new TextRun({
                text: "Senior .NET Developer | ASP.NET Core & FinTech Software Engineer",
                bold: true,
                font: "Calibri",
                size: 23, // 11.5pt
                color: primaryBlue,
              }),
            ],
          }),

          // CONTACT LINE
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 160 },
            children: [
              new TextRun({ text: "Karachi, Pakistan  •  +92 333 3442076  •  ", font: "Calibri", size: 19, color: grayText }),
              new ExternalHyperlink({
                children: [
                  new TextRun({ text: "bilalxdeveloper99@gmail.com", style: "Hyperlink", font: "Calibri", size: 19, color: primaryBlue, underline: {} }),
                ],
                link: "mailto:bilalxdeveloper99@gmail.com",
              }),
              new TextRun({ text: "  •  ", font: "Calibri", size: 19, color: grayText }),
              new ExternalHyperlink({
                children: [
                  new TextRun({ text: "LinkedIn", style: "Hyperlink", font: "Calibri", size: 19, color: primaryBlue, underline: {} }),
                ],
                link: "https://linkedin.com/in/bilal-ahmed-2572a4244",
              }),
              new TextRun({ text: "  •  ", font: "Calibri", size: 19, color: grayText }),
              new ExternalHyperlink({
                children: [
                  new TextRun({ text: "Live Portfolio", style: "Hyperlink", font: "Calibri", size: 19, color: primaryBlue, underline: {} }),
                ],
                link: "https://bkahmad25.github.io/MyPortfolio",
              }),
              new TextRun({ text: "  •  ", font: "Calibri", size: 19, color: grayText }),
              new ExternalHyperlink({
                children: [
                  new TextRun({ text: "GitHub", style: "Hyperlink", font: "Calibri", size: 19, color: primaryBlue, underline: {} }),
                ],
                link: "https://github.com/bkahmad25",
              }),
            ],
          }),

          // SECTION: PROFESSIONAL SUMMARY
          createSectionHeading("Professional Summary"),
          new Paragraph({
            spacing: { before: 40, after: 120 },
            alignment: AlignmentType.JUSTIFIED,
            children: [
              new TextRun({
                text: "Results-driven ",
                font: "Calibri",
                size: 20,
              }),
              new TextRun({
                text: "Senior .NET Developer & FinTech Software Engineer ",
                bold: true,
                font: "Calibri",
                size: 20,
              }),
              new TextRun({
                text: "with 3+ years of experience designing, modernizing, and scaling high-performance enterprise systems and stock brokerage back-office platforms. Specialized in ",
                font: "Calibri",
                size: 20,
              }),
              new TextRun({
                text: "ASP.NET Core (.NET 8), C#, Entity Framework Core, SignalR, Microservices, and high-throughput SQL Server ",
                bold: true,
                font: "Calibri",
                size: 20,
              }),
              new TextRun({
                text: "optimization (stored procedures, indexing, execution plans). Proven track record in engineering mission-critical financial calculation engines (FIFO Average Price, Capital Gains Tax, CDC/NCCPL settlements), real-time market watch feeds, and Azure cloud CI/CD deployments. Adept at Clean Architecture, SOLID design principles, and collaborative Agile/Scrum delivery.",
                font: "Calibri",
                size: 20,
              }),
            ],
          }),

          // SECTION: TECHNICAL SKILLS
          createSectionHeading("Technical Skills"),
          new Paragraph({
            spacing: { before: 40, after: 30 },
            children: [
              new TextRun({ text: "• Backend & Frameworks: ", bold: true, font: "Calibri", size: 20, color: darkText }),
              new TextRun({ text: "C#, ASP.NET Core (.NET 8/7/6), ASP.NET MVC, Web API, Web Forms, Entity Framework Core, LINQ, Dapper, SignalR, Microservices", font: "Calibri", size: 20 }),
            ],
          }),
          new Paragraph({
            spacing: { before: 30, after: 30 },
            children: [
              new TextRun({ text: "• Database & Tuning: ", bold: true, font: "Calibri", size: 20, color: darkText }),
              new TextRun({ text: "MS SQL Server, PostgreSQL, MySQL, Complex Stored Procedures, T-SQL, Query Optimization, Indexing, Execution Plans, Triggers", font: "Calibri", size: 20 }),
            ],
          }),
          new Paragraph({
            spacing: { before: 30, after: 30 },
            children: [
              new TextRun({ text: "• Frontend & Web: ", bold: true, font: "Calibri", size: 20, color: darkText }),
              new TextRun({ text: "Angular, React.js, JavaScript (ES6+), TypeScript, HTML5, CSS3, Bootstrap 5, AJAX, JSON", font: "Calibri", size: 20 }),
            ],
          }),
          new Paragraph({
            spacing: { before: 30, after: 30 },
            children: [
              new TextRun({ text: "• Architecture & APIs: ", bold: true, font: "Calibri", size: 20, color: darkText }),
              new TextRun({ text: "RESTful Web APIs, SOAP Web Services, JWT Authentication, Clean Architecture, Repository Pattern, SOLID Principles, DI", font: "Calibri", size: 20 }),
            ],
          }),
          new Paragraph({
            spacing: { before: 30, after: 120 },
            children: [
              new TextRun({ text: "• Cloud, Tools & DevOps: ", bold: true, font: "Calibri", size: 20, color: darkText }),
              new TextRun({ text: "Microsoft Azure (App Services), Git, GitHub, IIS Server, CI/CD Pipelines, Visual Studio 2022, SSMS, Crystal Reports", font: "Calibri", size: 20 }),
            ],
          }),

          // SECTION: PROFESSIONAL EXPERIENCE
          createSectionHeading("Professional Experience"),

          // Role 1: Catalyst IT Solutions
          createItemHeader("Software Developer (.NET Developer)", "Catalyst IT Solutions", "Mar 2025 – Present | Karachi, Pakistan"),
          createBullet("High-Availability FinTech Architecture: ", "Engineer and maintain high-volume financial brokerage back-office applications using ASP.NET Core 8, C#, and Entity Framework Core, maintaining 99.9% production uptime."),
          createBullet("Database Latency Reduction: ", "Refactored high-concurrency SQL Server stored procedures and executed indexing optimizations, slashing query execution latency by 35% on high-traffic financial ledgers."),
          createBullet("Real-Time SignalR Streaming: ", "Architected real-time market data feed module using ASP.NET Core SignalR, distributing sub-second live price updates to concurrent brokerage trader terminals."),
          createBullet("Financial & Payment Integrations: ", "Developed secure REST & SOAP APIs with JWT authentication connecting national clearing entities (CDC / NCCPL) and commercial banking gateways."),
          createBullet("Cloud & DevOps: ", "Collaborated with agile engineering teams to automate build and deployment workflows to Microsoft Azure via automated CI/CD pipelines."),

          // Role 2: Out2Sol IT Services
          createItemHeader("Associate Software Developer", "Out2Sol IT Services & Consulting", "Oct 2023 – Feb 2025 | Karachi, Pakistan"),
          createBullet("Enterprise Web Applications: ", "Engineered and enhanced line-of-business software platforms using ASP.NET MVC, C#, LINQ, and SQL Server for financial, healthcare, and insurance clients."),
          createBullet("Database Procedures & Analytics: ", "Authored 50+ optimized SQL Server stored procedures, triggers, and views for automated audit reporting, financial ledgers, and executive dashboards."),
          createBullet("Performance Optimization: ", "Identified and resolved transaction deadlocks and query bottlenecks, improving API throughput by 25% across high-volume modules."),
          createBullet("UI & Reporting: ", "Integrated interactive Angular and Bootstrap user interfaces; configured automated Crystal Reports for daily statutory compliance statements."),

          // Role 3: Alkaram Textile Mills
          createItemHeader("ERP Operations Specialist (Data Systems)", "Alkaram Textile Mills Pvt Ltd", "Apr 2022 – Aug 2023 | Karachi, Pakistan"),
          createBullet("Data Reconciliation: ", "Managed automated data migration and verification workflows from spreadsheet records to relational databases with 99.8% record accuracy."),
          createBullet("Inventory & Schema Auditing: ", "Monitored inventory lifecycle workflows, providing domain data schemas and operational requirements for ERP modernization."),

          // SECTION: KEY ENTERPRISE & FINTECH PROJECTS
          createSectionHeading("Key Enterprise & FinTech Projects"),

          createItemHeader("Financial Backoffice Management System (GEM)", "ASP.NET Core, MVC, SQL Server, REST/SOAP", "Brokerage Platform"),
          createBullet("Brokerage Operations: ", "Developed and maintained core modules for Order Execution, Client Portfolio Management, Risk Assessment, End-of-Day (EOD) Batch Processing, and CDC Clearing reconciliation."),
          createBullet("Automated Reporting: ", "Engineered automated Crystal Reports workflows for daily client settlements and regulatory audits, reducing manual reconciliation effort by 15+ hours weekly."),

          createItemHeader("FIFO Average Price & CGT Tax Calculation Engine", "C#, ASP.NET Core, SQL Server Stored Procedures", "FinTech Algorithm"),
          createBullet("Algorithmic Calculation: ", "Designed First-In-First-Out (FIFO) share inventory calculation engine to dynamically calculate real-time average buy prices across multiple order executions."),
          createBullet("Capital Gains Tax (CGT): ", "Automated realized and unrealized CGT tax calculations adhering strictly to stock exchange statutory regulations and financial audit rules."),

          createItemHeader("Real-Time Market Watch System", "ASP.NET Core SignalR, WebSockets, JavaScript, SQL Server", "Live Trading Tool"),
          createBullet("Tick Streaming: ", "Built low-latency push notification pipeline delivering live tick-by-tick stock quotes, volumes, and price fluctuations to multi-client brokerage dashboards."),

          createItemHeader("Crescent Steel ERP Management System", "ASP.NET Core, C#, MS SQL Server, REST API", "Enterprise ERP"),
          createBullet("Split Share Automation: ", "Engineered corporate actions module automating stock splits, instantaneously reallocating share quantities and adjusting shareholder ledgers with zero discrepancy."),

          createItemHeader("Digital Account Opening & Insurance Agent Portal", "ASP.NET Core, Web API, EF Core, SQL Server, Angular", "KYC & Onboarding"),
          createBullet("Digital Onboarding: ", "Automated customer KYC and digital onboarding workflow, reducing customer account opening turnaround time by 40%."),
          createBullet("Policy Lifecycle: ", "Implemented policy issuance workflows, automated debit/credit voucher generation, and sales agent commission distribution."),

          // SECTION: EDUCATION & CREDENTIALS
          createSectionHeading("Education & Credentials"),

          new Paragraph({
            spacing: { before: 80, after: 40 },
            children: [
              new TextRun({ text: "Bachelor of Science in Computer Science (BSCS)", bold: true, font: "Calibri", size: 21, color: darkText }),
              new TextRun({ text: " — Virtual University of Pakistan", bold: true, font: "Calibri", size: 21, color: primaryBlue }),
              new TextRun({ text: "    (2026 – In Progress)", italics: true, font: "Calibri", size: 19, color: grayText }),
            ],
          }),
          new Paragraph({
            spacing: { before: 40, after: 40 },
            children: [
              new TextRun({ text: "Higher Diploma in Software Engineering (HDSE)", bold: true, font: "Calibri", size: 21, color: darkText }),
              new TextRun({ text: " — Aptech Computer Education", bold: true, font: "Calibri", size: 21, color: primaryBlue }),
              new TextRun({ text: "    (2021 – 2024 | Grade A)", italics: true, font: "Calibri", size: 19, color: grayText }),
            ],
          }),
          new Paragraph({
            spacing: { before: 40, after: 60 },
            children: [
              new TextRun({ text: "Intermediate in Computer Science (ICS)", bold: true, font: "Calibri", size: 21, color: darkText }),
              new TextRun({ text: " — Modern Sciences College, Karachi", bold: true, font: "Calibri", size: 21, color: primaryBlue }),
              new TextRun({ text: "    (2019 – 2021 | Grade A)", italics: true, font: "Calibri", size: 19, color: grayText }),
            ],
          }),
        ],
      },
    ],
  });

  return doc;
}

async function main() {
  const doc = createATSResumeDocx();
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync('Bilal_Ahmed_Resume.docx', buffer);
  console.log('Successfully generated Bilal_Ahmed_Resume.docx. Size:', buffer.length, 'bytes');
}

main().catch(err => {
  console.error('Error generating docx:', err);
  process.exit(1);
});
