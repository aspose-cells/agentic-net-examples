// Title: Add a custom-sized, dark‑blue chart title to a column chart using Aspose.Cells for .NET
// AI Prompts: Create a column chart from worksheet data and set its title text, font size to 14 points, and color to DarkBlue with Aspose.Cells C#. | Modify an existing Excel chart's title to use a specific font size and color programmatically via the Aspose.Cells .NET API. | Generate a workbook, populate sales data, add a column chart, and apply custom font styling to the chart title in C#.
// Common Searches: how to change the font size and color of an Excel chart title using Aspose.Cells in C# | Aspose.Cells example for setting chart title style programmatically | C# code to add a dark blue title to a column chart in an Excel file | customize chart title appearance Aspose.Cells .NET tutorial
// Tags: Aspose.Cells chart title font styling | custom chart title color Aspose.Cells | column chart title customization .NET | Excel chart title formatting C# | programmatic chart title appearance Aspose.Cells

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Populate some sample data for the chart
sheet.Cells["A1"].PutValue("Month");
sheet.Cells["B1"].PutValue("Sales");
sheet.Cells["A2"].PutValue("Jan");
sheet.Cells["A3"].PutValue("Feb");
sheet.Cells["A4"].PutValue("Mar");
sheet.Cells["B2"].PutValue(1200);
sheet.Cells["B3"].PutValue(1500);
sheet.Cells["B4"].PutValue(1800);

// Add a column chart (you can change ChartType as needed)
int chartIndex = sheet.Charts.Add(Aspose.Cells.Charts.ChartType.Column, 5, 0, 20, 10);
Aspose.Cells.Charts.Chart chart = sheet.Charts[chartIndex];

// Set the data range for the chart series
chart.NSeries.Add("B2:B4", true);
chart.NSeries.CategoryData = "A2:A4";

// Insert a chart title and customize its appearance
chart.Title.Text = "Quarterly Sales";
chart.Title.Font.Size = 14;                     // Custom font size
chart.Title.Font.Color = System.Drawing.Color.DarkBlue; // Custom font color

// Save the workbook to a file
workbook.Save("ChartWithCustomTitle.xlsx");
