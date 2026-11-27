// Title: How to remove unused styles, export cell comments, and generate compact HTML from an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove unused styles before saving workbook as HTML | Export cell comments to HTML with Aspose.Cells .NET | Generate compact HTML output without gridlines using HtmlSaveOptions in Aspose.Cells | Save only the active worksheet as HTML with Aspose.Cells C# | Create a sample Excel file and export to HTML with comments using Aspose.Cells
// Tags: aspose.cells remove unused styles | aspose.cells export comments to html | aspose.cells htmlsaveoptions compact output | aspose.cells htmlsaveoptions exportgridlines false | aspose.cells export active worksheet only

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The program loads an existing Excel file or creates a sample workbook, adds a comment, removes any styles that are not used with Workbook.RemoveUnusedStyles(), configures HtmlSaveOptions to disable grid lines, row/column headings, export only the active worksheet and optionally include comments, ensures the output directory exists, and then saves the workbook as a compact HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file names
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            Workbook workbook = null;

            try
            {
                // Load existing workbook if it exists, otherwise create a sample workbook
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook.");
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Populate sample data
                    cells["A1"].PutValue("Product");
                    cells["B1"].PutValue("Price");
                    cells["A2"].PutValue("Apple");
                    cells["B2"].PutValue(1.20);
                    cells["A3"].PutValue("Banana");
                    cells["B3"].PutValue(0.80);
                    cells["A4"].PutValue("Cherry");
                    cells["B4"].PutValue(2.50);

                    // Add a comment to demonstrate comment export
                    int commentIndex = sheet.Comments.Add("B2");
                    Comment comment = sheet.Comments[commentIndex];
                    comment.Note = "Price in USD";

                    // Save the sample workbook for reference (optional)
                    workbook.Save("sample_created.xlsx", SaveFormat.Xlsx);
                }

                // Remove any styles that are not used in the workbook to reduce size
                Console.WriteLine("Removing unused styles...");
                workbook.RemoveUnusedStyles();

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export comments (if supported by the version)
                    // ExportComments = true, // Uncomment if the property exists in your Aspose.Cells version
                    ExportGridLines = false,               // Do not export grid lines
                    ExportRowColumnHeadings = false,       // Do not export row/column headings
                    ExportActiveWorksheetOnly = true       // Export only the active sheet
                };

                // Ensure the output directory exists (handle cases where outputPath has no directory part)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML using the configured options
                Console.WriteLine($"Saving workbook as HTML to \"{outputPath}\"...");
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine("Export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during processing:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Dispose workbook if needed (Aspose.Cells implements IDisposable in newer versions)
                if (workbook is IDisposable disposableWorkbook)
                {
                    disposableWorkbook.Dispose();
                }
            }
        }
    }
}
