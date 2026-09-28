// Title: Generate IE‑compatible HTML with sheet‑tab navigation from a multi‑sheet workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# program that creates a workbook with three worksheets, adds a column chart, saves it as HTML via Aspose.Cells, and launches the file in the default browser to manually test sheet‑tab switching in Internet Explorer. | Enhance the example to programmatically verify that the exported HTML contains the JavaScript required for tab navigation and that the file opens without error on Windows.
// Common Searches: c# aspnet export multi‑sheet workbook to html with tab navigation using aspose.cells | how to make Aspose.Cells HTML output work in Internet Explorer | save Excel workbook as html with sheet tabs for IE compatibility in .NET | verify sheet tab switching in Aspose.Cells generated html file | open generated html file automatically after saving with Aspose.Cells C#
// Tags: save workbook as html using Aspose.Cells | sheet tab navigation in Aspose.Cells HTML output | export column chart to html with Aspose.Cells | launch generated html file from C# application | internet explorer compatibility for Aspose.Cells HTML

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program builds a three‑sheet workbook (Summary, Data, Chart), adds a column chart, saves the workbook as an HTML file with built‑in sheet‑tab navigation using Aspose.Cells, and opens the file in the default browser so the user can confirm that clicking the tabs switches the displayed sheet in Internet Explorer.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with multiple worksheets
            Workbook workbook = new Workbook();

            // First sheet (default) – Summary
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Summary";
            sheet1.Cells["A1"].PutValue("This is the Summary sheet.");
            sheet1.Cells["A2"].PutValue(DateTime.Now);

            // Second sheet – Data
            Worksheet sheet2 = workbook.Worksheets.Add("Data");
            sheet2.Cells["A1"].PutValue("ID");
            sheet2.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 10; i++)
            {
                sheet2.Cells[$"A{i}"].PutValue(i - 1);
                sheet2.Cells[$"B{i}"].PutValue($"Item {i - 1}");
            }

            // Third sheet – Chart
            Worksheet sheet3 = workbook.Worksheets.Add("Chart");
            sheet3.Cells["A1"].PutValue("Category");
            sheet3.Cells["B1"].PutValue("Amount");
            sheet3.Cells["A2"].PutValue("A");
            sheet3.Cells["B2"].PutValue(120);
            sheet3.Cells["A3"].PutValue("B");
            sheet3.Cells["B3"].PutValue(80);
            sheet3.Cells["A4"].PutValue("C");
            sheet3.Cells["B4"].PutValue(150);

            // Create a column chart on the third sheet
            int chartIndex = sheet3.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet3.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Sample Column Chart";

            // Define the path for the HTML output (MHTML not supported in older versions)
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MultiSheetWorkbook.html");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            try
            {
                workbook.Save(outputPath, SaveFormat.Html);
                Console.WriteLine($"Workbook saved as HTML to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook as HTML: {ex.Message}");
                return;
            }

            // Launch the default browser to open the generated HTML file
            if (File.Exists(outputPath))
            {
                try
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo
                    {
                        FileName = outputPath,
                        UseShellExecute = true
                    };
                    Process.Start(startInfo);
                    Console.WriteLine("File opened in the default browser. Verify that clicking the sheet tabs switches the displayed sheet.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to open the file in the default browser.");
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("The HTML file was not found after saving.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
