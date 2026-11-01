// Title: Add a bold Arial 12‑point title to a column chart with Aspose.Cells for .NET
// AI Prompts: Create a new workbook, insert a column chart, and set its title to "Sample Label" using Arial 12‑pt bold font via Aspose.Cells C#. | Apply Arial 12‑point bold styling to the title of an existing column chart in an Aspose.Cells workbook using .NET.
// Common Searches: Aspose.Cells C# set chart title font to Arial 12 bold | How to format a chart title in Aspose.Cells .NET | Add a bold title to a column chart with Aspose.Cells | Change Excel chart title font programmatically using Aspose.Cells | Set chart title properties in C# Aspose.Cells API
// Tags: Aspose.Cells chart title font styling | set chart title Arial bold .NET | column chart label formatting Aspose.Cells | Excel chart title customization C# | Aspose.Cells workbook chart creation

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Add a chart (Column type) to the worksheet
int chartIndex = sheet.Charts.Add(Aspose.Cells.Charts.ChartType.Column, 5, 0, 15, 5);
Aspose.Cells.Charts.Chart chart = sheet.Charts[chartIndex];

// Set the label (chart title) text
chart.Title.Text = "Sample Label";

// Configure the label font: Arial, size 12, bold
chart.Title.Font.Name = "Arial";
chart.Title.Font.Size = 12;
chart.Title.Font.IsBold = true;

// Save the workbook to a file
workbook.Save("ChartWithLabel.xlsx");
