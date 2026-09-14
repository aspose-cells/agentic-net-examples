// Title: Create a progress bar chart in Aspose.Cells (C#) and save to XLSX while keeping cell formulas intact
// AI Prompts: Generate C# code with Aspose.Cells that builds a bar chart from formula‑calculated progress percentages and writes the workbook to an XLSX file while retaining every formula. | Update an existing Aspose.Cells workbook to add a green progress bar chart, assign task names as X‑axis categories, show percentages as data labels, and export the file without converting formulas to values.
// Common Searches: Aspose.Cells C# how to keep formulas when saving workbook to XLSX | create a progress bar chart from formula cells using Aspose.Cells .NET | set X axis categories to task names in Aspose.Cells chart | display percentage data labels on Aspose.Cells bar chart | export Aspose.Cells workbook with formulas unchanged
// Tags: Aspose.Cells create bar chart from formula cells | Aspose.Cells export to XLSX with formulas intact | Aspose.Cells set chart XValues range | Aspose.Cells format data labels as percentage | Aspose.Cells apply green fill to bar series

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The example creates a workbook, fills task data with a formula for progress percentage, adds a green bar chart linked to the progress column, configures task names as X‑axis categories, shows percentage data labels, and saves the file as an XLSX workbook while preserving all cell formulas.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Progress";

            // Header row
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Completed");
            sheet.Cells["C1"].PutValue("Total");
            sheet.Cells["D1"].PutValue("Progress %");

            // Sample data
            string[] tasks = { "Design", "Development", "Testing", "Deployment" };
            int[] completed = { 30, 50, 20, 5 };
            int[] total = { 100, 100, 100, 100 };

            // Fill data rows and set formula for progress percentage
            for (int i = 0; i < tasks.Length; i++)
            {
                int row = i + 2; // Data starts at row 2
                sheet.Cells[row, 0].PutValue(tasks[i]);          // Column A: Task name
                sheet.Cells[row, 1].PutValue(completed[i]);      // Column B: Completed
                sheet.Cells[row, 2].PutValue(total[i]);          // Column C: Total
                sheet.Cells[row, 3].Formula = $"=B{row}/C{row}"; // Column D: Progress %
            }

            // Add a bar chart to visualize progress
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 6, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Project Progress";

            // Define the data range for the series (Progress % values)
            int firstDataRow = 2;
            int lastDataRow = tasks.Length + 1;
            chart.NSeries.Add($"D{firstDataRow}:D{lastDataRow}", true);

            // Set category (X) axis to task names
            // In some Aspose.Cells versions the property is XValues; use it for compatibility
            chart.NSeries[0].XValues = $"A{firstDataRow}:A{lastDataRow}";
            chart.NSeries[0].Name = "Progress";

            // Format the series: green fill for the bars
            chart.NSeries[0].Area.ForegroundColor = Color.Green;

            // Show data labels as percentages
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[0].DataLabels.NumberFormat = "0%";

            // Save the workbook to XLSX while preserving all cell formulas
            workbook.Save("ProgressBarChart.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
