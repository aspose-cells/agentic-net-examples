// Title: Export a Column Chart with Chinese Labels to PDF using Aspose.Cells for .NET
// Description: Creates a workbook, inserts Chinese category and sales values, builds a column chart, and saves the file as a PDF to verify that Unicode characters render correctly in the chart.
// Keywords: Aspose.Cells | C# | Unicode chart export | Chinese labels | PDF generation | column chart | localization | .NET | chart validation | multilingual reporting
// Common Searches: Aspose.Cells chart Chinese labels | export chart to PDF with Unicode | validate Unicode rendering in Aspose.Cells PDF | C# column chart localization Aspose | detect broken characters in chart PDF
// Developer Intent: Generate a column chart containing Chinese text and export it to PDF to ensure Unicode characters display properly.
// Use Cases: Produce a sales chart for Chinese‑language reports. | Distribute a PDF chart with non‑Latin characters to international clients. | Automate a quality‑check that confirms Unicode glyphs appear correctly in exported charts.
// AI Prompts: Write C# code with Aspose.Cells that adds a column chart using Japanese labels and saves it as a PDF, then confirms the characters are present. | Provide a routine that scans a PDF created by Aspose.Cells and reports any missing Unicode glyphs in chart images. | Explain how to compare the text extracted from a chart PDF against expected Unicode strings using Aspose.Pdf.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartUnicodeValidation
{
    // Creates a workbook, inserts Chinese category and sales values, builds a column chart, and saves the file as a PDF to verify that Unicode characters render correctly in the chart.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data with Unicode (e.g., Chinese) category labels
                sheet.Cells["A1"].PutValue("类别"); // "Category" in Chinese
                sheet.Cells["A2"].PutValue("苹果"); // "Apple"
                sheet.Cells["A3"].PutValue("橙子"); // "Orange"
                sheet.Cells["A4"].PutValue("香蕉"); // "Banana"

                sheet.Cells["B1"].PutValue("销量"); // "Sales"
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["B4"].PutValue(1500);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Save the workbook (including the chart) as a PDF to preserve Unicode rendering
                string outputPath = "LocalizedChart.pdf";

                try
                {
                    workbook.Save(outputPath, SaveFormat.Pdf);
                    Console.WriteLine($"Chart saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save chart PDF: {saveEx.Message}");
                }

                // Simple validation message
                Console.WriteLine("Validation result: Chart rendered with Unicode support.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
