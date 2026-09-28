// Title: How to right‑justify data label text in a horizontal bar chart using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a horizontal bar chart and sets the data label TextHorizontalAlignment to Right. | Show how to enable data labels for the first series of a bar chart and apply right‑aligned text using the TextAlignmentType enum in Aspose.Cells. | Update an existing Aspose.Cells chart to change its data label alignment to right‑justified without altering other formatting.
// Common Searches: Aspose.Cells C# set data label alignment to right for bar chart | right align data labels in a horizontal bar chart using Aspose.Cells .NET | how to change TextHorizontalAlignment of chart data labels in Aspose.Cells | C# example for right‑justified data labels in a bar chart with Aspose.Cells
// Tags: Aspose.Cells chart data label right alignment | horizontal bar chart data labels C# | TextHorizontalAlignment property Aspose.Cells | NSeries data label formatting .NET | Aspose.Cells bar chart label justification

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells;

// Create a new workbook and get the first worksheet
Workbook workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

// Populate sample data for the horizontal bar chart
sheet.Cells["A1"].PutValue("Category");
sheet.Cells["A2"].PutValue("A");
sheet.Cells["A3"].PutValue("B");
sheet.Cells["A4"].PutValue("C");

sheet.Cells["B1"].PutValue("Value");
sheet.Cells["B2"].PutValue(10);
sheet.Cells["B3"].PutValue(20);
sheet.Cells["B4"].PutValue(30);

// Add a horizontal bar chart (ChartType.Bar) to the worksheet
int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 15, 5);
Chart chart = sheet.Charts[chartIndex];

// Set the data range for the chart
chart.NSeries.Add("B2:B4", true);          // Values
chart.NSeries.CategoryData = "A2:A4";      // Categories

// Enable data labels for the first series
chart.NSeries[0].DataLabels.ShowValue = true;

// Right‑justify the text inside the data labels
chart.NSeries[0].DataLabels.TextHorizontalAlignment = TextAlignmentType.Right;

// Save the workbook
workbook.Save("HorizontalBarChart_WithRightJustifiedDataLabels.xlsx");
