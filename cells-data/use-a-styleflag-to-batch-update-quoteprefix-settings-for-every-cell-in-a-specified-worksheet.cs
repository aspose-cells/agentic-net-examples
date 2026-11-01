// Title: Batch enable QuotePrefix for all cells in a worksheet using Aspose.Cells StyleFlag in C#
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, creates a style with QuotePrefix = true, configures a StyleFlag to apply only QuotePrefix, and applies it to every cell in the first worksheet. | Write a C# snippet that uses Aspose.Cells to set the QuotePrefix flag on an entire worksheet without altering existing cell formatting. | Provide a complete C# console program that checks for an input workbook, creates it if missing, applies a QuotePrefix‑only style to all cells via StyleFlag, and saves the result.
// Common Searches: aspnet cells set quote prefix for entire sheet | apply styleflag quote prefix to all cells c# | batch update quote prefix attribute in excel using aspose.cells | preserve existing formatting while enabling quote prefix aspose.cells | how to use StyleFlag to change only QuotePrefix in Aspose.Cells
// Tags: Aspose.Cells StyleFlag QuotePrefix | C# batch apply cell style | set Excel QuotePrefix programmatically | apply style without changing other formatting | load and save workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixExample
{
    // The example loads an existing workbook (or creates a new one), defines a style with QuotePrefix enabled, configures a StyleFlag so only the QuotePrefix property is affected, applies this style to every cell in the first worksheet, and saves the updated workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = @"C:\Path\To\Input.xlsx";
            string outputPath = @"C:\Path\To\Output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file not found at '{inputPath}'. A new workbook will be created.");
                    workbook = new Workbook(); // creates a workbook with a default worksheet
                }

                Worksheet worksheet;

                // Get the first worksheet if it exists; otherwise add a new worksheet and retrieve it
                if (workbook.Worksheets.Count > 0)
                {
                    worksheet = workbook.Worksheets[0];
                }
                else
                {
                    int newIndex = workbook.Worksheets.Add();
                    worksheet = workbook.Worksheets[newIndex];
                }

                // Create a style with QuotePrefix enabled
                Style quoteStyle = workbook.CreateStyle();
                quoteStyle.QuotePrefix = true;

                // Configure a StyleFlag to apply only the QuotePrefix attribute
                StyleFlag flag = new StyleFlag
                {
                    QuotePrefix = true
                };

                // Apply the style to all cells in the worksheet
                worksheet.Cells.ApplyStyle(quoteStyle, flag);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
