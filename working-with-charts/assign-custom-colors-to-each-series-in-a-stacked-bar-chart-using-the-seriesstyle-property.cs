using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Create a new workbook and get the first worksheet
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];

// Populate sample data for a stacked bar chart
worksheet.Cells["A1"].PutValue("Category");
worksheet.Cells["A2"].PutValue("Q1");
worksheet.Cells["A3"].PutValue("Q2");
worksheet.Cells["A4"].PutValue("Q3");

worksheet.Cells["B1"].PutValue("Series1");
worksheet.Cells["B2"].PutValue(10);
worksheet.Cells["B3"].PutValue(20);
worksheet.Cells["B4"].PutValue(30);

worksheet.Cells["C1"].PutValue("Series2");
worksheet.Cells["C2"].PutValue(15);
worksheet.Cells["C3"].PutValue(25);
worksheet.Cells["C4"].PutValue(35);

worksheet.Cells["D1"].PutValue("Series3");
worksheet.Cells["D2"].PutValue(12);
worksheet.Cells["D3"].PutValue(22);
worksheet.Cells["D4"].PutValue(32);

// Add a stacked bar chart to the worksheet
int chartIndex = worksheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 15);
Chart chart = worksheet.Charts[chartIndex];

// Set the data range for the series (all three series) and the category axis
chart.NSeries.Add("B2:D4", true);
chart.NSeries.CategoryData = "A2:A4";

// Assign custom colors to each series using the Area.ForegroundColor property
chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189);   // Color for Series1
chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(192, 80, 77);   // Color for Series2
chart.NSeries[2].Area.ForegroundColor = Color.FromArgb(155, 187, 89);  // Color for Series3

// Save the workbook
workbook.Save("StackedBarCustomColors.xlsx");