// Title: C# Aspose.Cells example – verify chart title and axis labels display localized text from worksheet cells
// Common Searches: how to set an Excel chart title from a localized cell value using Aspose.Cells in C# | asp osee cells verify that axis labels use Spanish text from worksheet cells | example of reading cell strings for chart titles and axis titles with Aspose.Cells .NET | C# code to compare chart title text with source cell value in Aspose.Cells | localize Excel chart titles and axis labels with Aspose.Cells column chart
// Tags: chart title from worksheet cell Aspose.Cells | axis label localization Aspose.Cells | verify chart text matches cell value C# | localized Excel chart generation Aspose.Cells | column chart with cell‑based titles Aspose.Cells | Spanish chart titles Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, adds numeric data, writes Spanish strings into cells for the chart title and axis titles, inserts a column chart, assigns the chart's Title, CategoryAxis.Title, and ValueAxis.Title from those cells, checks that each title matches the corresponding cell value, outputs the verification results, and saves the file as ChartLocalization.xlsx.
class ChartLocalizationVerification
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data for the chart
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["A4"].PutValue(30);
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(35);

            // Insert localized strings into cells (e.g., Spanish)
            sheet.Cells["D1"].PutValue("Ventas Mensuales"); // Chart title
            sheet.Cells["E1"].PutValue("Mes");             // X‑axis title
            sheet.Cells["F1"].PutValue("Valor");           // Y‑axis title

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the series
            chart.NSeries.Add("A2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set chart title and axis titles from the localized cells
            chart.Title.Text = sheet.Cells["D1"].StringValue;
            chart.CategoryAxis.Title.Text = sheet.Cells["E1"].StringValue;
            chart.ValueAxis.Title.Text = sheet.Cells["F1"].StringValue;

            // Verify that the titles displayed on the chart match the source cell values
            bool isChartTitleCorrect = chart.Title.Text == sheet.Cells["D1"].StringValue;
            bool isXAxisTitleCorrect = chart.CategoryAxis.Title.Text == sheet.Cells["E1"].StringValue;
            bool isYAxisTitleCorrect = chart.ValueAxis.Title.Text == sheet.Cells["F1"].StringValue;

            Console.WriteLine("Chart title verification: " + isChartTitleCorrect);
            Console.WriteLine("X‑axis title verification: " + isXAxisTitleCorrect);
            Console.WriteLine("Y‑axis title verification: " + isYAxisTitleCorrect);

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = "ChartLocalization.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
