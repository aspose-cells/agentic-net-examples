// Title: Set custom number formats for PivotChart axes with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, accesses the first chart (assumed to be a PivotChart), applies a currency format to the value‑axis tick labels and a date format to the category‑axis tick labels, then saves the modified file.
// Keywords: Aspose.Cells chart axis format C# | custom number format PivotChart .NET | value axis currency format Aspose.Cells | category axis date format C# | set tick label format Aspose.Cells | Excel chart formatting Aspose.Cells | pivot chart axis formatting
// Common Searches: Aspose.Cells set value axis number format | How to format PivotChart axis labels in C# | Apply date format to chart category axis using Aspose.Cells | Change currency display on chart axis Aspose.Cells .NET | Save workbook after modifying chart axis format Aspose.Cells
// Developer Intent: Load an existing Excel file, customize the number formats of a PivotChart’s axes, and save the updated workbook.
// Use Cases: Present sales figures with a currency symbol on the value axis of a financial PivotChart. | Show project milestones using a readable date format on the category axis of a timeline chart. | Standardize axis formatting across multiple charts before exporting the workbook to PDF or image.
// AI Prompts: Generate C# code with Aspose.Cells that applies a custom currency format to the value axis and a custom date format to the category axis of a PivotChart, then saves the workbook. | Show how to set a percentage format on a chart’s value axis and a short‑date format on its category axis using Aspose.Cells for .NET. | Explain how to loop through all charts in a worksheet and assign the same number format to each axis tick label with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, accesses the first chart (assumed to be a PivotChart), applies a currency format to the value‑axis tick labels and a date format to the category‑axis tick labels, then saves the modified file.
class Program
{
    static void Main()
    {
        // Load the workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one chart (the PivotChart)
        if (sheet.Charts.Count > 0)
        {
            // Access the first chart
            Chart chart = sheet.Charts[0];

            // Apply a custom number format to the value axis tick labels
            chart.ValueAxis.TickLabels.NumberFormat = "$#,##0.00";

            // Optionally, apply a custom format to the category (X) axis tick labels
            chart.CategoryAxis.TickLabels.NumberFormat = "mmm dd, yyyy";
        }

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
