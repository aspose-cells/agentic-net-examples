// Title: Save a Progress Bar Chart to XLSX without Calculating Formulas – Aspose.Cells C# Example
// Description: Creates a new Workbook, inserts task data, builds a stacked bar progress chart, disables automatic formula calculation, and saves the file as XLSX while keeping all cell formulas unchanged.
// Keywords: Aspose.Cells C# save workbook | preserve formulas on save | stacked bar progress chart | ChartType.BarStacked example | disable CalculateOnSave | export chart to XLSX | Aspose.Cells formula settings
// Common Searches: how to stop formula calculation when saving Aspose.Cells workbook | Aspose.Cells progress bar chart C# example | save workbook with chart without recalculating formulas | ChartType.BarStacked Aspose.Cells tutorial | preserve cell formulas in exported XLSX
// Developer Intent: Generate a stacked‑bar progress chart, keep existing formulas intact, and export the workbook to XLSX using Aspose.Cells for .NET.
// Use Cases: Project‑status dashboards where task completion percentages are shown as progress bars and formulas must remain editable after export. | Financial models that include stacked bar visuals but require formulas to stay unchanged for downstream calculations. | Template‑driven reports that embed charts and need to be saved as XLSX without triggering automatic recalculation.
// AI Prompts: Write C# code with Aspose.Cells to add a stacked bar progress chart, turn off CalculateOnSave, and save as XLSX. | Explain the effect of Workbook.Settings.FormulaSettings.CalculateOnSave and how to use it to preserve formulas when saving a chart‑containing workbook. | Show how to make the progress bar data range dynamic (e.g., using named ranges) while still preventing formula evaluation on save.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart and ChartType

namespace ProgressBarChartSaveExample
{
    // Creates a new Workbook, inserts task data, builds a stacked bar progress chart, disables automatic formula calculation, and saves the file as XLSX while keeping all cell formulas unchanged.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data for a progress bar chart (using a stacked bar)
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Completed");
                sheet.Cells["C1"].PutValue("Remaining");

                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(30);   // Completed percentage
                sheet.Cells["C2"].PutValue(70);   // Remaining percentage

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(55);
                sheet.Cells["C3"].PutValue(45);

                // Add a stacked bar chart to represent the progress bar
                // Note: In Aspose.Cells the enum value is BarStacked (not StackedBar)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Progress Bar";

                // Series for completed portion
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries[0].Name = "Completed";

                // Series for remaining portion
                chart.NSeries.Add("C2:C3", true);
                chart.NSeries[1].Name = "Remaining";

                // Set category (task names)
                chart.NSeries.CategoryData = "A2:A3";

                // Ensure formulas are preserved (no calculation on save)
                workbook.Settings.FormulaSettings.CalculateOnSave = false;

                // Save the workbook as XLSX while preserving formulas
                workbook.Save("ProgressBarChart.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
