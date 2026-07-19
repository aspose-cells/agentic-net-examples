// Title: C# Example: Translate Excel Chart Category Axis Labels with Aspose.Cells and an External Service
// Description: This Aspose.Cells for .NET sample creates a workbook, adds a column chart, reads the category‑axis texts, translates each label via a placeholder service, writes the translated strings back to the source cells, recalculates the chart, and saves the result as an XLSX file.
// Keywords: Aspose.Cells C# chart axis translation | update Excel chart category labels programmatically | external translation API Aspose.Cells | localize Excel chart axis text | C# Aspose.Cells chart manipulation
// Common Searches: how to change chart axis labels with Aspose.Cells .NET | translate Excel chart categories using Aspose.Cells | update chart axis after modifying worksheet cells C# | Aspose.Cells example for multilingual chart labels | C# code to localize Excel chart axis
// Developer Intent: Replace the original category‑axis labels of an Excel chart with translated text and save the updated workbook.
// Use Cases: Generate multilingual reports by translating chart axis labels before distribution. | Dynamically adjust axis texts based on runtime data, such as adding language suffixes. | Integrate a third‑party translation API to produce localized Excel charts for global audiences.
// AI Prompts: Write C# code that calls Google Translate API to replace the placeholder Translate method and updates chart axis labels in Aspose.Cells. | Show how to batch‑translate axis labels for multiple charts in the same workbook using Aspose.Cells. | Add comprehensive error handling and logging around the translation call while updating chart axis texts.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET sample creates a workbook, adds a column chart, reads the category‑axis texts, translates each label via a placeholder service, writes the translated strings back to the source cells, recalculates the chart, and saves the result as an XLSX file.
    public class TranslateChartAxisLabels
    {
        // Placeholder for an external translation service.
        // In a real scenario, replace this with an actual API call.
        private static string Translate(string text)
        {
            // Example translation logic (append suffix)
            return text + "_translated";
        }

        public static void Run()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Populate sample data ----------
                // Category (axis labels) in column A
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["A4"].PutValue("Cherry");

                // Values in column B
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(80);
                worksheet.Cells["B4"].PutValue(150);

                // ---------- Add a column chart ----------
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set data range for series and categories
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Category labels (axis)

                // Calculate the chart to generate initial axis labels
                chart.Calculate();

                // ---------- Retrieve original axis labels ----------
                Axis categoryAxis = chart.CategoryAxis;
                string[] originalLabels = categoryAxis.GetAxisTexts();

                // ---------- Translate each label ----------
                string[] translatedLabels = new string[originalLabels.Length];
                for (int i = 0; i < originalLabels.Length; i++)
                {
                    translatedLabels[i] = Translate(originalLabels[i]);
                }

                // ---------- Write translated labels back to the worksheet ----------
                // Assuming the category data range starts at A2 and has the same length as originalLabels
                for (int i = 0; i < translatedLabels.Length; i++)
                {
                    // Row index: 2 + i (since A2 is the first label)
                    int row = 2 + i;
                    worksheet.Cells[row, 0].PutValue(translatedLabels[i]); // Column 0 = A
                }

                // Recalculate the chart to reflect updated labels
                chart.Calculate();

                // ---------- Save the workbook ----------
                string outputPath = "TranslatedAxisLabelsChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TranslateChartAxisLabels.Run();
        }
    }
}
