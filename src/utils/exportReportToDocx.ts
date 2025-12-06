import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  HeadingLevel, 
  AlignmentType, 
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  ShadingType,
  convertInchesToTwip,
  PageOrientation,
} from 'docx';
import { saveAs } from 'file-saver';
import { format } from 'date-fns';

interface ReportData {
  dateRange: string;
  executiveKPIs: {
    totalApplications: number;
    approvedApplications: number;
    pendingApplications: number;
    rejectedApplications: number;
    totalRevenue: number;
    collectedRevenue: number;
    pendingRevenue: number;
    totalInspections: number;
    completedInspections: number;
    scheduledInspections: number;
    avgComplianceScore: number;
    approvalRate: number;
    collectionRate: number;
    totalEntities: number;
    activeEntities: number;
    totalProjectValue: number;
    totalIntents: number;
  };
  investmentByLevel: {
    level2: { value: number; count: number };
    level3: { value: number; count: number };
    total: number;
    totalCount: number;
  };
  investmentYear: number;
  provincialData: Array<{ province: string; activePermits: number }>;
  permitTypeDistribution: Array<{ name: string; value: number }>;
  statusDistribution: Array<{ name: string; value: number }>;
  entityTypeDistribution: Array<{ name: string; value: number }>;
  complianceTabData: {
    totalPermits: number;
    totalComplianceReports: number;
    permitsBySector: Array<{ sector: string; count: number }>;
    totalInspectionsCarried: number;
    violationsReported: number;
  };
  monthlyTrends: Array<{ month: string; applications: number; approvals: number; revenue: number }>;
}

const formatCurrency = (amount: number) => {
  return new Intl.NumberFormat('en-PG', { style: 'currency', currency: 'PGK', maximumFractionDigits: 0 }).format(amount);
};

// Create a simple table for KPI cards
const createKPITable = (items: Array<{ label: string; value: string | number }>) => {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: items.map(item => 
          new TableCell({
            width: { size: Math.floor(100 / items.length), type: WidthType.PERCENTAGE },
            shading: { fill: "E8F5E9", type: ShadingType.CLEAR },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: item.label, bold: true, size: 20 }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: String(item.value), bold: true, size: 28 }),
                ],
              }),
            ],
          })
        ),
      }),
    ],
  });
};

// Create a data table
const createDataTable = (headers: string[], rows: string[][]) => {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      // Header row
      new TableRow({
        tableHeader: true,
        children: headers.map(header => 
          new TableCell({
            shading: { fill: "1B5E20", type: ShadingType.CLEAR },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: header, bold: true, color: "FFFFFF", size: 20 }),
                ],
              }),
            ],
          })
        ),
      }),
      // Data rows
      ...rows.map((row, rowIndex) => 
        new TableRow({
          children: row.map(cell => 
            new TableCell({
              shading: { fill: rowIndex % 2 === 0 ? "FFFFFF" : "F5F5F5", type: ShadingType.CLEAR },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: cell, size: 20 }),
                  ],
                }),
              ],
            })
          ),
        })
      ),
    ],
  });
};

// Create a visual horizontal bar chart using table cells with colored fills
const createHorizontalBarChart = (
  title: string,
  data: Array<{ label: string; value: number; color?: string }>,
  maxValue?: number
) => {
  // Return just a title with "No data" message if data is empty
  if (!data || data.length === 0) {
    return [
      new Paragraph({
        spacing: { before: 200, after: 100 },
        children: [
          new TextRun({ text: title, bold: true, size: 22, color: "1B5E20" }),
        ],
      }),
      new Paragraph({
        spacing: { before: 50, after: 100 },
        children: [
          new TextRun({ text: "No data available", italics: true, size: 18, color: "666666" }),
        ],
      }),
    ];
  }
  
  const max = maxValue || Math.max(...data.map(d => d.value), 1);
  const barColors = ["4CAF50", "2196F3", "FF9800", "9C27B0", "00BCD4", "E91E63"];
  
  const chartRows = data.map((item, index) => {
    const barWidth = Math.round((item.value / max) * 100);
    const color = item.color || barColors[index % barColors.length];
    
    return new TableRow({
      children: [
        // Label column
        new TableCell({
          width: { size: 30, type: WidthType.PERCENTAGE },
          verticalAlign: "center",
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { before: 50, after: 50 },
              children: [
                new TextRun({ text: item.label, size: 18 }),
              ],
            }),
          ],
        }),
        // Bar column
        new TableCell({
          width: { size: 50, type: WidthType.PERCENTAGE },
          children: [
            new Paragraph({
              spacing: { before: 50, after: 50 },
              children: [
                new TextRun({ 
                  text: "█".repeat(Math.max(1, Math.round(barWidth / 5))), 
                  color: color,
                  size: 20,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        // Value column
        new TableCell({
          width: { size: 20, type: WidthType.PERCENTAGE },
          verticalAlign: "center",
          children: [
            new Paragraph({
              alignment: AlignmentType.LEFT,
              spacing: { before: 50, after: 50 },
              children: [
                new TextRun({ text: ` ${item.value}`, bold: true, size: 18 }),
              ],
            }),
          ],
        }),
      ],
    });
  });

  return [
    new Paragraph({
      spacing: { before: 200, after: 100 },
      children: [
        new TextRun({ text: title, bold: true, size: 22, color: "1B5E20" }),
      ],
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.NONE },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
        insideHorizontal: { style: BorderStyle.NONE },
        insideVertical: { style: BorderStyle.NONE },
      },
      rows: chartRows,
    }),
  ];
};

// Create a visual pie chart legend with colored blocks
const createPieChartLegend = (
  title: string,
  data: Array<{ name: string; value: number; percentage: number }>
) => {
  // Return just a title with "No data" message if data is empty
  if (!data || data.length === 0) {
    return [
      new Paragraph({
        spacing: { before: 200, after: 100 },
        children: [
          new TextRun({ text: title, bold: true, size: 22, color: "1B5E20" }),
        ],
      }),
      new Paragraph({
        spacing: { before: 50, after: 100 },
        children: [
          new TextRun({ text: "No data available", italics: true, size: 18, color: "666666" }),
        ],
      }),
    ];
  }

  const colors = ["1B5E20", "2196F3", "FF9800", "9C27B0", "00BCD4", "E91E63"];
  
  const legendItems = data.map((item, index) => {
    const color = colors[index % colors.length];
    return new TableRow({
      children: [
        // Color block
        new TableCell({
          width: { size: 10, type: WidthType.PERCENTAGE },
          shading: { fill: color, type: ShadingType.CLEAR },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: " ", size: 20 })],
            }),
          ],
        }),
        // Label
        new TableCell({
          width: { size: 50, type: WidthType.PERCENTAGE },
          children: [
            new Paragraph({
              spacing: { before: 30, after: 30 },
              children: [
                new TextRun({ text: `  ${item.name}`, size: 20 }),
              ],
            }),
          ],
        }),
        // Value
        new TableCell({
          width: { size: 20, type: WidthType.PERCENTAGE },
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { before: 30, after: 30 },
              children: [
                new TextRun({ text: String(item.value), bold: true, size: 20 }),
              ],
            }),
          ],
        }),
        // Percentage
        new TableCell({
          width: { size: 20, type: WidthType.PERCENTAGE },
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { before: 30, after: 30 },
              children: [
                new TextRun({ text: `${item.percentage.toFixed(1)}%`, size: 20, color: "666666" }),
              ],
            }),
          ],
        }),
      ],
    });
  });

  // Create ASCII pie representation
  const pieSymbols = ["●", "◐", "◑", "○", "◒", "◓"];
  const pieRepresentation = data.slice(0, 6).map((item, i) => 
    `${pieSymbols[i]} ${item.name}: ${item.percentage.toFixed(0)}%`
  ).join("   ");

  return [
    new Paragraph({
      spacing: { before: 200, after: 100 },
      children: [
        new TextRun({ text: title, bold: true, size: 22, color: "1B5E20" }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 150 },
      children: [
        new TextRun({ text: pieRepresentation, size: 18, color: "333333" }),
      ],
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: legendItems,
    }),
  ];
};

// Create section heading
const createSectionHeading = (text: string) => {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 200 },
    children: [
      new TextRun({
        text: text,
        bold: true,
        size: 28,
        color: "1B5E20",
      }),
    ],
  });
};

// Create subsection heading
const createSubsectionHeading = (text: string) => {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 150 },
    children: [
      new TextRun({
        text: text,
        bold: true,
        size: 24,
      }),
    ],
  });
};

// Create description paragraph with placeholder
const createDescriptionPlaceholder = (context: string) => {
  return new Paragraph({
    spacing: { before: 100, after: 200 },
    children: [
      new TextRun({
        text: `[Description: ${context}]`,
        italics: true,
        color: "666666",
        size: 20,
      }),
    ],
  });
};

// Create a horizontal line separator
const createSeparator = () => {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: {
      bottom: { color: "1B5E20", style: BorderStyle.SINGLE, size: 6 },
    },
    children: [],
  });
};

export const exportReportToDocx = async (data: ReportData) => {
  // Fetch the PNG emblem as base64
  let emblemImage: ImageRun | null = null;
  try {
    const response = await fetch('/images/png-emblem.png');
    const blob = await response.blob();
    const arrayBuffer = await blob.arrayBuffer();
    emblemImage = new ImageRun({
      data: arrayBuffer,
      transformation: { width: 80, height: 80 },
    });
  } catch (error) {
    console.error('Failed to load emblem image:', error);
  }

  const headerElements: Paragraph[] = [];

  // Add emblem if loaded
  if (emblemImage) {
    headerElements.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
        children: [emblemImage],
      })
    );
  }

  // Organization Header - matching the official format
  headerElements.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 150 },
      children: [
        new TextRun({
          text: "Conservation & Environment Protection Authority",
          bold: true,
          size: 28,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 50 },
      children: [
        new TextRun({
          text: "PERMIT MANAGEMENT REPORT",
          bold: true,
          size: 32,
          color: "1B5E20",
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 50 },
      children: [
        new TextRun({
          text: `Report Generated: ${format(new Date(), 'dd MMMM yyyy, HH:mm')}`,
          size: 20,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 150 },
      children: [
        new TextRun({
          text: `Data Period: ${data.dateRange}`,
          size: 20,
        }),
      ],
    }),
    createSeparator()
  );

  // Calculate totals for percentages
  const totalStatusCount = data.statusDistribution.reduce((sum, item) => sum + item.value, 0);
  const totalEntityCount = data.entityTypeDistribution.reduce((sum, item) => sum + item.value, 0);
  const totalPermitTypeCount = data.permitTypeDistribution.reduce((sum, item) => sum + item.value, 0);

  // Prepare pie chart data with percentages
  const statusPieData = data.statusDistribution.map(item => ({
    name: item.name,
    value: item.value,
    percentage: totalStatusCount > 0 ? (item.value / totalStatusCount) * 100 : 0,
  }));

  const entityTypePieData = data.entityTypeDistribution.map(item => ({
    name: item.name,
    value: item.value,
    percentage: totalEntityCount > 0 ? (item.value / totalEntityCount) * 100 : 0,
  }));

  // Executive Overview Section
  const executiveOverviewSection = [
    createSectionHeading("1. Executive Overview"),
    createDescriptionPlaceholder("Summary of overall permit management performance and key metrics"),
    createSubsectionHeading("Key Performance Indicators"),
    createKPITable([
      { label: "Total Applications", value: data.executiveKPIs.totalApplications },
      { label: "Approval Rate", value: `${data.executiveKPIs.approvalRate}%` },
      { label: "Revenue Collected", value: formatCurrency(data.executiveKPIs.collectedRevenue) },
      { label: "Compliance Score", value: `${data.executiveKPIs.avgComplianceScore}%` },
    ]),
    new Paragraph({ spacing: { before: 200 } }),
    createKPITable([
      { label: "Registered Entities", value: data.executiveKPIs.totalEntities },
      { label: "Active Entities", value: data.executiveKPIs.activeEntities },
      { label: "Completed Inspections", value: data.executiveKPIs.completedInspections },
      { label: "Pending Applications", value: data.executiveKPIs.pendingApplications },
    ]),
    new Paragraph({ spacing: { before: 300 } }),
    
    // Application Status Distribution - PIE CHART
    ...createPieChartLegend(
      "📊 Application Status Distribution",
      statusPieData
    ),
    
    new Paragraph({ spacing: { before: 300 } }),
    
    // Permit Type Distribution - BAR CHART
    ...createHorizontalBarChart(
      "📈 Top Permit Types by Volume",
      data.permitTypeDistribution.slice(0, 8).map(item => ({
        label: item.name.length > 25 ? item.name.substring(0, 22) + "..." : item.name,
        value: item.value,
      }))
    ),
    
    new Paragraph({ spacing: { before: 300 } }),
    
    // Entity Type Distribution - PIE CHART
    ...createPieChartLegend(
      "🏢 Entity Type Distribution",
      entityTypePieData
    ),
  ];

  // Investment Value Section
  const investmentSection = [
    createSectionHeading("2. Investment Value Analysis"),
    createDescriptionPlaceholder("Analysis of project investment values by activity level"),
    createSubsectionHeading(`Investment Summary for ${data.investmentYear}`),
    
    // Investment by Level - BAR CHART
    ...createHorizontalBarChart(
      "📊 Investment Value by Activity Level",
      [
        { label: "Level 2 Activities", value: Math.round(data.investmentByLevel.level2.value / 1000000), color: "2196F3" },
        { label: "Level 3 Activities", value: Math.round(data.investmentByLevel.level3.value / 1000000), color: "4CAF50" },
      ]
    ),
    new Paragraph({
      spacing: { before: 50, after: 100 },
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "(Values in Millions PGK)", size: 16, color: "888888", italics: true }),
      ],
    }),
    
    new Paragraph({ spacing: { before: 200 } }),
    createDataTable(
      ["Activity Level", "Investment Value", "Number of Permits"],
      [
        ["Level 2 Activities", formatCurrency(data.investmentByLevel.level2.value), String(data.investmentByLevel.level2.count)],
        ["Level 3 Activities", formatCurrency(data.investmentByLevel.level3.value), String(data.investmentByLevel.level3.count)],
        ["Total", formatCurrency(data.investmentByLevel.total), String(data.investmentByLevel.totalCount)],
      ]
    ),
  ];

  // Geographic Analysis Section - BAR CHART
  const activeProvinces = data.provincialData.filter(p => p.activePermits > 0).slice(0, 10);
  const geographicSection = [
    createSectionHeading("3. Geographic Analysis"),
    createDescriptionPlaceholder("Provincial distribution of active permits across Papua New Guinea"),
    
    // Provincial Distribution - BAR CHART
    ...createHorizontalBarChart(
      "📍 Active Permits by Province (Top 10)",
      activeProvinces.map(item => ({
        label: item.province.replace(' Province', '').substring(0, 20),
        value: item.activePermits,
      }))
    ),
    
    new Paragraph({ spacing: { before: 300 } }),
    createSubsectionHeading("Complete Provincial Breakdown"),
    createDataTable(
      ["Province", "Active Permits"],
      data.provincialData.filter(p => p.activePermits > 0).map(item => [
        item.province,
        String(item.activePermits),
      ])
    ),
  ];

  // Compliance & Enforcement Section - PIE + BAR
  const totalSectorCount = data.complianceTabData.permitsBySector.reduce((sum, item) => sum + item.count, 0);
  const sectorPieData = data.complianceTabData.permitsBySector.slice(0, 6).map(item => ({
    name: item.sector,
    value: item.count,
    percentage: totalSectorCount > 0 ? (item.count / totalSectorCount) * 100 : 0,
  }));

  const complianceSection = [
    createSectionHeading("4. Compliance & Enforcement"),
    createDescriptionPlaceholder("Overview of compliance monitoring and enforcement activities"),
    createSubsectionHeading("Compliance Statistics"),
    createKPITable([
      { label: "Total Permits", value: data.complianceTabData.totalPermits },
      { label: "Compliance Reports", value: data.complianceTabData.totalComplianceReports },
      { label: "Inspections Carried Out", value: data.complianceTabData.totalInspectionsCarried },
      { label: "Violations Reported", value: data.complianceTabData.violationsReported },
    ]),
    
    new Paragraph({ spacing: { before: 300 } }),
    
    // Permits by Sector - PIE CHART
    ...createPieChartLegend(
      "🏭 Permits by Sector/Industry",
      sectorPieData
    ),
    
    new Paragraph({ spacing: { before: 300 } }),
    
    // Sector Bar Chart
    ...createHorizontalBarChart(
      "📊 Sector Volume Comparison",
      data.complianceTabData.permitsBySector.slice(0, 8).map(item => ({
        label: item.sector.substring(0, 20),
        value: item.count,
      }))
    ),
  ];

  // Financial Performance Section - with charts
  const revenueData = [
    { name: "Collected", value: data.executiveKPIs.collectedRevenue, percentage: data.executiveKPIs.totalRevenue > 0 ? (data.executiveKPIs.collectedRevenue / data.executiveKPIs.totalRevenue) * 100 : 0 },
    { name: "Outstanding", value: data.executiveKPIs.pendingRevenue, percentage: data.executiveKPIs.totalRevenue > 0 ? (data.executiveKPIs.pendingRevenue / data.executiveKPIs.totalRevenue) * 100 : 0 },
  ];

  const financialSection = [
    createSectionHeading("5. Financial Performance"),
    createDescriptionPlaceholder("Revenue collection and financial performance metrics"),
    createSubsectionHeading("Revenue Summary"),
    createKPITable([
      { label: "Total Revenue", value: formatCurrency(data.executiveKPIs.totalRevenue) },
      { label: "Collected", value: formatCurrency(data.executiveKPIs.collectedRevenue) },
      { label: "Outstanding", value: formatCurrency(data.executiveKPIs.pendingRevenue) },
      { label: "Collection Rate", value: `${data.executiveKPIs.collectionRate}%` },
    ]),
    
    new Paragraph({ spacing: { before: 300 } }),
    
    // Revenue Status - PIE CHART
    ...createPieChartLegend(
      "💰 Revenue Collection Status",
      revenueData
    ),
    
    new Paragraph({ spacing: { before: 300 } }),
    
    // Monthly Revenue Trend - BAR CHART
    ...createHorizontalBarChart(
      "📈 Monthly Revenue Trend (Last 6 Months)",
      data.monthlyTrends.slice(-6).map(item => ({
        label: item.month,
        value: Math.round(item.revenue / 1000),
      }))
    ),
    new Paragraph({
      spacing: { before: 50, after: 100 },
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "(Values in Thousands PGK)", size: 16, color: "888888", italics: true }),
      ],
    }),
    
    new Paragraph({ spacing: { before: 200 } }),
    createSubsectionHeading("Monthly Revenue Detail"),
    createDataTable(
      ["Month", "Revenue Collected"],
      data.monthlyTrends.slice(-6).map(item => [
        item.month,
        formatCurrency(item.revenue),
      ])
    ),
  ];

  // Trends & Forecasting Section - with charts
  const trendsSection = [
    createSectionHeading("6. Trends & Forecasting"),
    createDescriptionPlaceholder("Historical trends and forecasting analysis for strategic planning"),
    
    // Applications vs Approvals - BAR CHART
    ...createHorizontalBarChart(
      "📊 Monthly Applications Volume",
      data.monthlyTrends.slice(-6).map(item => ({
        label: item.month,
        value: item.applications,
        color: "2196F3",
      }))
    ),
    
    new Paragraph({ spacing: { before: 200 } }),
    
    ...createHorizontalBarChart(
      "✅ Monthly Approvals Volume",
      data.monthlyTrends.slice(-6).map(item => ({
        label: item.month,
        value: item.approvals,
        color: "4CAF50",
      }))
    ),
    
    new Paragraph({ spacing: { before: 300 } }),
    createSubsectionHeading("Monthly Application & Approval Details"),
    createDataTable(
      ["Month", "Applications", "Approvals", "Revenue"],
      data.monthlyTrends.slice(-6).map(item => [
        item.month,
        String(item.applications),
        String(item.approvals),
        formatCurrency(item.revenue),
      ])
    ),
    new Paragraph({ spacing: { before: 300 } }),
    createSubsectionHeading("Executive Insights"),
    new Paragraph({
      spacing: { before: 100, after: 100 },
      children: [
        new TextRun({
          text: "✅ Positive Indicators:",
          bold: true,
          size: 22,
          color: "1B5E20",
        }),
      ],
    }),
    new Paragraph({
      bullet: { level: 0 },
      children: [
        new TextRun({
          text: `${data.executiveKPIs.approvalRate}% approval rate demonstrates effective regulatory framework`,
          size: 20,
        }),
      ],
    }),
    new Paragraph({
      bullet: { level: 0 },
      children: [
        new TextRun({
          text: `${formatCurrency(data.executiveKPIs.collectedRevenue)} revenue collected supports government initiatives`,
          size: 20,
        }),
      ],
    }),
    new Paragraph({
      bullet: { level: 0 },
      children: [
        new TextRun({
          text: `${data.executiveKPIs.totalEntities} registered entities contributing to economic development`,
          size: 20,
        }),
      ],
    }),
    new Paragraph({
      spacing: { before: 200, after: 100 },
      children: [
        new TextRun({
          text: "⚠️ Areas Requiring Attention:",
          bold: true,
          size: 22,
          color: "FF6F00",
        }),
      ],
    }),
    new Paragraph({
      bullet: { level: 0 },
      children: [
        new TextRun({
          text: `${data.executiveKPIs.pendingApplications} applications pending review require expedited processing`,
          size: 20,
        }),
      ],
    }),
    new Paragraph({
      bullet: { level: 0 },
      children: [
        new TextRun({
          text: `${formatCurrency(data.executiveKPIs.pendingRevenue)} in outstanding payments require follow-up`,
          size: 20,
        }),
      ],
    }),
  ];

  // Footer
  const footerSection = [
    createSeparator(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200 },
      children: [
        new TextRun({
          text: "--- End of Report ---",
          italics: true,
          size: 20,
          color: "666666",
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100 },
      children: [
        new TextRun({
          text: "Conservation & Environment Protection Authority © " + new Date().getFullYear(),
          size: 18,
          color: "999999",
        }),
      ],
    }),
  ];

  // Create the document
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: {
            width: convertInchesToTwip(8.27), // A4 width
            height: convertInchesToTwip(11.69), // A4 height
          },
          margin: {
            top: convertInchesToTwip(0.75),
            bottom: convertInchesToTwip(0.75),
            left: convertInchesToTwip(0.75),
            right: convertInchesToTwip(0.75),
          },
        },
      },
      children: [
        ...headerElements,
        ...executiveOverviewSection,
        ...investmentSection,
        ...geographicSection,
        ...complianceSection,
        ...financialSection,
        ...trendsSection,
        ...footerSection,
      ],
    }],
  });

  // Generate and save the document
  const blob = await Packer.toBlob(doc);
  const filename = `CEPA_Permit_Management_Report_${format(new Date(), 'yyyy-MM-dd_HHmm')}.docx`;
  saveAs(blob, filename);
  
  return filename;
};
