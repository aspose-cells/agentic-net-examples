// Title: C# method to load an Excel workbook and export HTML with tooltips using Aspose.Cells
// Description: A static utility that validates a file path, loads an Excel workbook via Aspose.Cells, enables HtmlSaveOptions.AddTooltipText, saves the workbook to a memory stream, and returns the resulting HTML string. Includes robust error handling.
// Keywords: Aspose.Cells | C# HTML export | AddTooltipText | Excel to HTML conversion | memory stream HTML | tooltip in exported HTML | HtmlSaveOptions | load workbook | export workbook as HTML | convert Excel to HTML C#
// Common Searches: Aspose.Cells export Excel to HTML with tooltips | C# convert Excel file to HTML string | Enable tooltip text in Aspose.Cells HTML export | Save Excel as HTML in memory stream Aspose | Get HTML from workbook without writing file Aspose.Cells
// Developer Intent: Create a reusable function that reads an Excel file and returns its HTML representation with cell comments rendered as tooltips.
// Use Cases: Show an HTML preview of an uploaded Excel file in a web portal while preserving comments as hover tooltips. | Generate email‑ready HTML from an Excel report without creating temporary files on disk. | Render Excel data with interactive tooltips in a WinForms/WPF application by injecting the returned HTML directly.
// AI Prompts: Write a unit test for LoadWorkbookAndExportHtml that asserts tooltip markup appears in the output HTML. | Refactor the method to accept a Stream and optional HtmlSaveOptions, enabling custom export settings. | Demonstrate how to embed the HTML string returned by LoadWorkbookAndExportHtml into an ASP.NET MVC view.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // A static utility that validates a file path, loads an Excel workbook via Aspose.Cells, enables HtmlSaveOptions.AddTooltipText, saves the workbook to a memory stream, and returns the resulting HTML string. Includes robust error handling.
    public static class HtmlUtility
    {
        /// <param name="excelFilePath">Full path to the source Excel file.</param>
        /// <returns>HTML representation of the workbook with tooltip text enabled.</returns>
        public static string LoadWorkbookAndExportHtml(string excelFilePath)
        {
            if (string.IsNullOrWhiteSpace(excelFilePath))
                throw new ArgumentException("Excel file path must be provided.", nameof(excelFilePath));

            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException("The specified Excel file was not found.", excelFilePath);

            try
            {
                // Load the workbook from the file.
                Workbook workbook = new Workbook(excelFilePath);

                // Configure HTML save options and enable tooltip text.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    AddTooltipText = true
                };

                // Save the workbook to a memory stream as HTML.
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, saveOptions);
                    htmlStream.Position = 0; // Reset stream position for reading.

                    // Read the HTML content from the stream and return it.
                    using (StreamReader reader = new StreamReader(htmlStream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                // Wrap and rethrow to preserve stack trace while providing context.
                throw new InvalidOperationException("Failed to convert Excel to HTML.", ex);
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                string excelPath;

                if (args.Length > 0)
                {
                    excelPath = args[0];
                }
                else
                {
                    Console.Write("Enter full path to the Excel file: ");
                    excelPath = Console.ReadLine();
                }

                // Validate the input path before processing.
                if (string.IsNullOrWhiteSpace(excelPath))
                    throw new ArgumentException("No Excel file path was provided.");

                string htmlContent = HtmlUtility.LoadWorkbookAndExportHtml(excelPath);

                // Output the HTML to console (or you could write to a file).
                Console.WriteLine("=== Generated HTML ===");
                Console.WriteLine(htmlContent);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File error: {fnfEx.Message}");
            }
            catch (ArgumentException argEx)
            {
                Console.Error.WriteLine($"Argument error: {argEx.Message}");
            }
            catch (InvalidOperationException invOpEx)
            {
                Console.Error.WriteLine($"Processing error: {invOpEx.Message}");
                Console.Error.WriteLine(invOpEx.InnerException?.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
