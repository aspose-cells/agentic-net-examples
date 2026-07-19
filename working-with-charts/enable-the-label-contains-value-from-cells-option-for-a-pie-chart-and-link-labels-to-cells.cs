// Title: Aspose.Cells for .NET – Link Pie Chart Data Labels to Cells (Label Contains – Value From Cells)
// Description: C# example that creates a workbook, adds categories, values and custom label text, inserts a pie chart, and configures the series to show values, category names, and to pull each data label from a cell range via the LinkedSource property. The sample also links the number format and prevents label overlap before saving the file.
// Keywords: Aspose.Cells C# pie chart data labels | LinkedSource property Aspose.Cells | Label Contains Value From Cells | prevent overlapping chart labels | custom label text from cells | Aspose.Cells for .NET chart example | pie chart label formatting Aspose
// Common Searches: Aspose.Cells link pie chart labels to cells | How to use Label Contains – Value From Cells in Aspose | C# set data label linked source Aspose.Cells | Avoid overlapping data labels in Aspose pie chart | Enable custom text for chart labels Aspose.Cells
// Developer Intent: Create a pie chart whose data labels are sourced from worksheet cells and automatically stay non‑overlapping.
// Use Cases: Sales dashboard where each slice shows quantity with units taken from a separate column. | Financial report chart that displays currency strings from cells, preserving the source number format. | Dynamic KPI pie chart that updates label text instantly when the underlying cells change.
// AI Prompts: Generate C# code with Aspose.Cells to add a pie chart and set DataLabels.LinkedSource to a cell range for custom label text. | Explain how ShowValue, ShowCategoryName, and IsNeverOverlap affect pie chart labels in Aspose.Cells. | Show how to refresh a chart after modifying the linked label cells in an existing workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsPieChartLabelFromCells
{
    // C# example that creates a workbook, adds categories, values and custom label text, inserts a pie chart, and configures the series to show values, category names, and to pull each data label from a cell range via the LinkedSource property. The sample also links the number format and prevents label overlap before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Cells that contain the custom label text (e.g., with units)
            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("30 pcs");
            sheet.Cells["C3"].PutValue("45 pcs");
            sheet.Cells["C4"].PutValue("25 pcs");

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B4", true);
            // Set the category (slice names)
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Enable data labels
            series.DataLabels.ShowValue = true;               // Show the numeric value
            series.DataLabels.ShowCategoryName = true;        // Show the category name
            // Link the label text to cells C2:C4 (Label Contains – Value From Cells)
            series.DataLabels.LinkedSource = "C2:C4";
            // Optionally link the number format so it follows the source cells
            series.DataLabels.NumberFormatLinked = true;

            // Optional: avoid overlapping labels for better readability
            series.DataLabels.IsNeverOverlap = true;

            // Save the workbook
            workbook.Save("PieChartLabelFromCells.xlsx");
        }
    }
}
