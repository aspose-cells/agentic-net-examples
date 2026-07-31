// Title: Aspose.Cells C# – Build a Progress‑Bar Style Stacked Bar Chart
// Description: Shows how to create an Excel workbook with a stacked bar chart that mimics a progress bar using Aspose.Cells for .NET. The sample adds task names and percentages, removes gaps (GapWidth = 0), fully overlaps bars (Overlap = 100), applies a green fill, and saves the file as .xlsx.
// Keywords: Aspose.Cells | C# | stacked bar chart | progress bar visualization | GapWidth | Overlap | series fill color | Excel chart automation | .NET chart example | task completion percentage
// Common Searches: Aspose.Cells progress bar chart C# | stacked bar without gaps Aspose | set overlap 100 Aspose.Cells | change series fill color Excel chart | export task progress to Excel using Aspose
// Developer Intent: Generate an Excel file that displays task completion as a solid‑filled progress indicator via a single stacked bar series.
// Use Cases: Automated status reports where each task’s completion percentage appears as a visual bar. | Dashboard worksheets that represent milestones with colored progress indicators. | Exporting KPI progress data to Excel for quick stakeholder review without manual formatting. | Creating printable project summaries that include compact progress visuals.
// AI Prompts: Write C# code with Aspose.Cells to add a stacked bar chart that functions as a progress bar, setting GapWidth to 0 and Overlap to 100. | Explain how to apply a custom fill color to the visible series in an Aspose.Cells stacked bar chart for progress representation. | Provide step‑by‑step instructions to bind category labels and percentage values to a progress‑style chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace ProgressBarChartDemo
{
    // Shows how to create an Excel workbook with a stacked bar chart that mimics a progress bar using Aspose.Cells for .NET. The sample adds task names and percentages, removes gaps (GapWidth = 0), fully overlaps bars (Overlap = 100), applies a green fill, and saves the file as .xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Sample data -----
            // Column A – Category (Task name)
            // Column B – Progress value (percentage)
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Progress");
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(70);   // 70%
            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(45);   // 45%
            sheet.Cells["A4"].PutValue("Task 3");
            sheet.Cells["B4"].PutValue(90);   // 90%

            // ----- Add a stacked bar chart -----
            // Use the Add method that takes (ChartType, topRow, leftColumn, bottomRow, rightColumn)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 22, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Configure the single visible series to look like a progress bar
            Series series = chart.NSeries[0];
            series.GapWidth = 0;       // Remove gaps between bars for a solid look
            series.Overlap = 100;      // Ensure bars are fully stacked (no spacing)

            // Optional: set a distinct fill color for the progress bar
            series.Area.ForegroundColor = Color.Green;

            // Save the workbook to an XLSX file
            workbook.Save("ProgressBarChart.xlsx");
        }
    }
}
