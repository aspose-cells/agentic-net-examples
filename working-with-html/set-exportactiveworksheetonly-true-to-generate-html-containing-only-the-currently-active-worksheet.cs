// Title: How to export only the active worksheet to an HTML file with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export active worksheet to HTML using HtmlSaveOptions | C# generate HTML from a single worksheet with Aspose.Cells | Set ExportActiveWorksheetOnly true in Aspose.Cells to save only the current sheet as HTML | Save specific sheet as HTML in .NET using Aspose.Cells HtmlSaveOptions | Create HTML output for the active sheet only with Aspose.Cells for .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | C# export single worksheet to HTML | active sheet HTML conversion Aspose.Cells | save specific worksheet as HTML .NET | Aspose.Cells generate HTML for active sheet

using System;
using System.IO;
using Aspose.Cells;

namespace ExportActiveWorksheetOnlyDemo
{
    // The example loads or creates a workbook, sets a chosen worksheet as active, configures HtmlSaveOptions with ExportActiveWorksheetOnly = true, and saves the workbook as an HTML file that contains only the active worksheet.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string inputPath = "SampleWorkbook.xlsx";
            string outputPath = "ActiveWorksheetOnly.html";

            Workbook workbook = null;

            try
            {
                // ---------------------------------------------------------------
                // 1. Prepare a sample workbook (if the input file does not exist)
                // ---------------------------------------------------------------
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook...");

                    workbook = new Workbook();
                    // Create three worksheets with some data
                    for (int i = 0; i < 3; i++)
                    {
                        Worksheet ws = workbook.Worksheets[i];
                        ws.Name = $"Sheet{i + 1}";
                        ws.Cells["A1"].PutValue($"This is {ws.Name}");
                        ws.Cells["A2"].PutValue(DateTime.Now);
                    }

                    // Save the sample workbook for future runs
                    workbook.Save(inputPath);
                    Console.WriteLine($"Sample workbook saved to \"{inputPath}\".");
                }
                else
                {
                    // Load the existing workbook
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }

                // ---------------------------------------------------------------
                // 2. Set the active worksheet (the one that will be exported)
                // ---------------------------------------------------------------
                // For demonstration, make the second worksheet active (index = 1)
                int activeSheetIndex = 1;
                if (activeSheetIndex >= workbook.Worksheets.Count)
                {
                    Console.WriteLine("Requested active sheet index is out of range. Using the first sheet instead.");
                    activeSheetIndex = 0;
                }

                workbook.Worksheets.ActiveSheetIndex = activeSheetIndex;
                Worksheet activeWorksheet = workbook.Worksheets[activeSheetIndex];
                Console.WriteLine($"Active worksheet set to \"{activeWorksheet.Name}\" (index {activeSheetIndex}).");

                // ---------------------------------------------------------------
                // 3. Configure HtmlSaveOptions to export only the active sheet
                // ---------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                htmlOptions.ExportActiveWorksheetOnly = true;   // key setting
                htmlOptions.ExportGridLines = true;             // optional visual preference

                // ---------------------------------------------------------------
                // 4. Save the workbook as HTML using the configured options
                // ---------------------------------------------------------------
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"HTML file generated successfully: \"{outputPath}\"");
                Console.WriteLine("Only the active worksheet has been exported, as requested.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }

            // ---------------------------------------------------------------
            // 5. Inform the user where to find the result
            // ---------------------------------------------------------------
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
