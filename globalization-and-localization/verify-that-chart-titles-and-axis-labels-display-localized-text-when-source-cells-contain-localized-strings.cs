// Title: Verify that Aspose.Cells C# chart titles and axis labels display localized text from worksheet cells
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart, reads English and Spanish strings from specific worksheet cells, assigns them to the chart Title, CategoryAxis.Title, and ValueAxis.Title, then prints boolean results confirming each assignment matches the source cell. | Write a C# example that populates localized strings in cells, binds those cells to a chart's title and axis labels using Aspose.Cells, validates the bindings, and saves the workbook.
// Common Searches: Aspose.Cells C# verify chart title matches cell value | How to bind Excel chart axis title to a localized cell using Aspose.Cells | C# Aspose.Cells chart localization test for English and Spanish labels | Check if Aspose.Cells chart titles update when cell text changes | Read localized strings from worksheet and apply to Aspose.Cells chart titles
// Tags: Aspose.Cells chart title from cell | Aspose.Cells localized chart axis labels | C# verify chart text matches worksheet cells | Aspose.Cells column chart localization | Excel chart title binding Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, writes English and Spanish strings into cells, adds a column chart, assigns the chart title and X/Y axis titles from those cells, verifies that each chart text matches the corresponding cell value, prints the verification results, and saves the workbook as LocalizedChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);

            // Localized strings placed in cells (e.g., English and Spanish)
            sheet.Cells["D1"].PutValue("Sales Overview");   // English chart title
            sheet.Cells["D2"].PutValue("Resumen de Ventas"); // Spanish chart title (alternative)
            sheet.Cells["E1"].PutValue("Months");           // English X‑axis label
            sheet.Cells["E2"].PutValue("Meses");            // Spanish X‑axis label (alternative)
            sheet.Cells["F1"].PutValue("Amount");           // English Y‑axis label
            sheet.Cells["F2"].PutValue("Cantidad");         // Spanish Y‑axis label (alternative)

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series data directly using a range address
            chart.NSeries.Add("B2:B3", true);

            // Assign chart title and axis titles from the localized cells (using English version here)
            chart.Title.Text = sheet.Cells["D1"].StringValue;               // Chart title
            chart.CategoryAxis.Title.Text = sheet.Cells["E1"].StringValue; // X‑axis title
            chart.ValueAxis.Title.Text = sheet.Cells["F1"].StringValue;    // Y‑axis title

            // Verification: ensure the chart titles and axis labels match the source cells
            bool titleMatches = chart.Title.Text == sheet.Cells["D1"].StringValue;
            bool xAxisMatches = chart.CategoryAxis.Title.Text == sheet.Cells["E1"].StringValue;
            bool yAxisMatches = chart.ValueAxis.Title.Text == sheet.Cells["F1"].StringValue;

            Console.WriteLine($"Chart title matches source cell: {titleMatches}");
            Console.WriteLine($"X‑axis title matches source cell: {xAxisMatches}");
            Console.WriteLine($"Y‑axis title matches source cell: {yAxisMatches}");

            // Save the workbook (optional, demonstrates persistence)
            string outputPath = "LocalizedChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
