// Title: How to add and display a title for a column chart with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add title to column chart and make it visible | set chart.Title.Text and Title.IsVisible in Aspose.Cells workbook | Aspose.Cells column chart title not showing how to enable | programmatically assign chart title in Aspose.Cells .NET example | Aspose.Cells chart title property usage in C#
// Tags: assign chart title text Aspose.Cells | enable chart title visibility Aspose.Cells | column chart title configuration Aspose.Cells | Aspose.Cells chart.Title.Text usage | Aspose.Cells chart.Title.IsVisible usage

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Get the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Add a column chart to the worksheet (position: row 5‑15, column 0‑5)
int chartIndex = sheet.Charts.Add(Aspose.Cells.Charts.ChartType.Column, 5, 0, 15, 5);
Aspose.Cells.Charts.Chart chart = sheet.Charts[chartIndex];

// Set the chart title text and make it visible
chart.Title.Text = "Sales Report";
chart.Title.IsVisible = true;

// Save the workbook to a file
workbook.Save("ChartTitle.xlsx");
