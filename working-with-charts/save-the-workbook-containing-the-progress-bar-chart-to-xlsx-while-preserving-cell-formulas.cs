// Title: Save XLSX with a Progress‑Bar Chart while Preserving Formulas using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, fills columns A and B with task labels and percentage values, adds a Bar chart that mimics a progress bar, disables formula recalculation by setting CalculateOnSave = false, and saves the file as XLSX so the original cell formulas remain unchanged.
// Keywords: Aspose.Cells C# save workbook | preserve formulas Aspose.Cells | progress bar chart Aspose.Cells | CalculateOnSave false | export chart to XLSX | .NET chart example | Aspose.Cells Bar chart
// Common Searches: save Aspose.Cells workbook without recalculating formulas | create progress bar chart in Aspose.Cells C# | disable CalculateOnSave Aspose.Cells .NET | export Aspose.Cells chart to XLSX | keep formulas when saving Excel with Aspose.Cells
// Developer Intent: Generate an XLSX file that contains a progress‑bar chart and retains the original cell formulas.
// Use Cases: Produce a project‑status report with a visual progress bar that can be updated later because formulas stay intact. | Automate batch creation of task‑tracking workbooks for a team, each with a progress chart that preserves calculation logic. | Expose a web API that returns a ready‑to‑download XLSX containing a progress bar, without triggering formula recalculation on the server.
// AI Prompts: Show how to apply custom colors to the progress‑bar chart while keeping CalculateOnSave disabled. | Provide code to open an existing workbook, modify the progress values, and save without recalculating formulas. | Explain how to export the progress‑bar chart as an image and embed it in the same workbook without affecting formulas.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsProgressBarChart
{
    // Creates a new Workbook, fills columns A and B with task labels and percentage values, adds a Bar chart that mimics a progress bar, disables formula recalculation by setting CalculateOnSave = false, and saves the file as XLSX so the original cell formulas remain unchanged.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data that will be used by the progress‑bar chart
                // Column A – Labels, Column B – Values (percentage)
                cells["A1"].PutValue("Task 1");
                cells["A2"].PutValue("Task 2");
                cells["A3"].PutValue("Task 3");
                cells["B1"].PutValue(0.4); // 40 %
                cells["B2"].PutValue(0.7); // 70 %
                cells["B3"].PutValue(0.9); // 90 %

                // Add a bar chart that will act as a progress bar
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data series (values) and category (labels)
                chart.NSeries.Add("B1:B3", true);
                chart.NSeries.CategoryData = "A1:A3";

                // Optional: format the chart to look like a progress bar
                chart.Title.Text = "Progress Bar";
                chart.ShowDataTable = false;               // corrected property
                chart.Legend.Position = LegendPositionType.Bottom;

                // Ensure formulas are not recalculated on save (preserve original formulas)
                workbook.Settings.FormulaSettings.CalculateOnSave = false;

                // Save the workbook to XLSX format (lifecycle: save)
                string outputPath = "ProgressBarChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
