// Title: Aspose.Cells for .NET – Verify Chart Title and Axis Labels Use Localized Cell Values (Japanese)
// Description: This C# example creates a workbook, writes Japanese strings into cells, builds a column chart, assigns those cell values to the chart’s Title, CategoryAxis.Title, and ValueAxis.Title, optionally applies SettableChartGlobalizationSettings, and saves the file so you can confirm that the chart displays localized titles and axis labels.
// Keywords: Aspose.Cells | C# | chart localization | Excel chart title from cell | Japanese chart titles | axis label localization | SettableChartGlobalizationSettings | GlobalizationSettings | internationalization | Excel .NET | Aspose.Cells chart API
// Common Searches: Aspose.Cells chart title from cell | display Japanese text in Excel chart using Aspose.Cells | set axis labels from worksheet cells C# | globalization settings for Aspose.Cells charts | verify localized chart labels Aspose.Cells | how to use SettableChartGlobalizationSettings .NET | Aspose.Cells localization example
// Developer Intent: Confirm that chart titles and axis labels reflect the localized strings stored in worksheet cells.
// Use Cases: Read Japanese strings from worksheet cells and assign them to chart.Title.Text, CategoryAxis.Title.Text, and ValueAxis.Title.Text. | Apply SettableChartGlobalizationSettings to customize default UI strings while preserving cell‑derived titles. | Generate an Excel file with a column chart that shows localized titles for sales, month, and amount. | Use the example as a template for other languages and chart types.
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart whose title and axis labels are taken from cells containing French text. | Show how to configure GlobalizationSettings to replace default chart UI strings without affecting titles sourced from cells. | Explain steps to programmatically verify that localized chart titles and axis labels appear correctly after saving the workbook. | Provide a PowerShell script that runs the compiled example and checks the chart titles in the resulting XLSX file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, writes Japanese strings into cells, builds a column chart, assigns those cell values to the chart’s Title, CategoryAxis.Title, and ValueAxis.Title, optionally applies SettableChartGlobalizationSettings, and saves the file so you can confirm that the chart displays localized titles and axis labels.
    public class VerifyChartLocalization
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Populate cells with localized strings (Japanese in this case)
                // ------------------------------------------------------------
                sheet.Cells["A1"].PutValue("売上");          // Chart title (Sales)
                sheet.Cells["A2"].PutValue("月");           // Category axis title (Month)
                sheet.Cells["A3"].PutValue("金額");         // Value axis title (Amount)

                // Sample data for the chart
                sheet.Cells["B1"].PutValue("Jan");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["C1"].PutValue("Feb");
                sheet.Cells["C2"].PutValue(1500);
                sheet.Cells["D1"].PutValue("Mar");
                sheet.Cells["D2"].PutValue(1800);

                // ------------------------------------------------------------
                // Add a column chart
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data series (values) and categories (months)
                chart.NSeries.Add("B2:D2", true);          // Values
                chart.NSeries.CategoryData = "B1:D1";      // Categories

                // ------------------------------------------------------------
                // Assign localized titles directly from the cells
                // ------------------------------------------------------------
                chart.Title.Text = sheet.Cells["A1"].StringValue;               // "売上"
                chart.CategoryAxis.Title.Text = sheet.Cells["A2"].StringValue; // "月"
                chart.ValueAxis.Title.Text = sheet.Cells["A3"].StringValue;    // "金額"

                // ------------------------------------------------------------
                // (Optional) Demonstrate using SettableChartGlobalizationSettings
                // to customize default UI strings, not required for cell‑based titles
                // ------------------------------------------------------------
                SettableChartGlobalizationSettings customSettings = new SettableChartGlobalizationSettings();
                customSettings.SetChartTitleName("カスタムタイトル"); // Custom chart title label
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = customSettings
                };

                // ------------------------------------------------------------
                // Save the workbook – the chart will display the localized titles
                // ------------------------------------------------------------
                workbook.Save("VerifyChartLocalization.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VerifyChartLocalization.Run();
        }
    }
}
