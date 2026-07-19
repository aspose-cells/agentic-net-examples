// Title: Add a 3‑D Column Chart to an Aspose.Cells Worksheet (C#)
// Description: Creates a new workbook, writes quarterly sales data to A1:B5, inserts a Column3D chart positioned from row 6 col 0 to row 20 col 10, binds the chart to the data range with headers, sets a custom title and built‑in style, then saves the file as an XLSX document.
// Keywords: Aspose.Cells 3D column chart C# | ChartType.Column3D example | add chart Aspose.Cells .NET | set chart data range Aspose.Cells | customize chart title Aspose.Cells | Excel dashboard C# Aspose | GitHub Aspose.Cells chart sample | save workbook with chart Aspose
// Common Searches: how to insert a 3D column chart using Aspose.Cells for .NET | Aspose.Cells C# set data range for Column3D chart | example of customizing title and style of a 3D column chart in Aspose.Cells | saving an Excel file that contains a 3D column chart with Aspose.Cells | Aspose.Cells chart positioning by cell coordinates
// Developer Intent: Insert a three‑dimensional column chart into an existing worksheet and persist the workbook.
// Use Cases: Generate a quarterly sales report with a visual 3D column chart for executive review. | Build an Excel‑based dashboard that compares product performance across regions using a Column3D chart. | Automate creation of presentation‑ready spreadsheets that include styled 3D charts for data analysis.
// AI Prompts: Provide C# code that adds a 3D column chart to a worksheet with Aspose.Cells, binds it to a data range, and sets a custom title and style. | Show how to position a Column3D chart at a specific cell range and apply a built‑in style in Aspose.Cells for .NET. | Explain how to change series colors, axis labels, and other visual properties of a 3D column chart created with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Assume a workbook has already been created or loaded earlier
Workbook workbook = new Workbook();                 // Create a new workbook (replace with load if needed)
Worksheet worksheet = workbook.Worksheets[0];       // Get the first worksheet

// Sample data for the chart
worksheet.Cells["A1"].PutValue("Category");
worksheet.Cells["A2"].PutValue("Q1");
worksheet.Cells["A3"].PutValue("Q2");
worksheet.Cells["A4"].PutValue("Q3");
worksheet.Cells["A5"].PutValue("Q4");

worksheet.Cells["B1"].PutValue("Sales");
worksheet.Cells["B2"].PutValue(1200);
worksheet.Cells["B3"].PutValue(1500);
worksheet.Cells["B4"].PutValue(1800);
worksheet.Cells["B5"].PutValue(2100);

// Add a three‑dimensional column chart (Column3D) to the worksheet
int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 6, 0, 20, 10);
Chart chart3D = worksheet.Charts[chartIndex];

// Define the data range for the chart (including headers)
chart3D.SetChartDataRange("A1:B5", true);

// Optional: customize appearance (e.g., title, style)
chart3D.Title.Text = "3D Column Chart";
chart3D.Style = 2; // Built‑in style index

// Save the workbook (replace with your desired path and format)
workbook.Save("ThreeDimensionalColumnChart.xlsx", SaveFormat.Xlsx);
