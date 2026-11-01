// Title: Create a C# method to convert an Excel workbook to an HTML string with tooltips using Aspose.Cells
// AI Prompts: Write a C# function that accepts a file path, loads the workbook with Aspose.Cells, sets HtmlSaveOptions.AddTooltipText to true, and returns the generated HTML as a UTF‑8 string. | Implement error handling that checks for file existence and wraps conversion errors when exporting an Excel file to HTML with tooltip support. | Show how to use a MemoryStream to save the workbook as HTML in Aspose.Cells without writing intermediate files.
// Common Searches: how to export an Excel file to HTML with cell comments as tooltips using Aspose.Cells in C# | Aspose.Cells HtmlSaveOptions AddTooltipText example code | convert workbook to HTML string without creating a temporary file .NET | C# load Excel workbook and get HTML output with tooltip text enabled | Aspose.Cells HTML export memory stream C# tutorial
// Tags: Aspose.Cells HTML conversion with tooltips | C# HtmlSaveOptions AddTooltipText | Export Excel to HTML string .NET | MemoryStream HTML generation Aspose.Cells | Workbook to HTML with tooltip support | Exception handling for Excel to HTML conversion

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Provides a C# utility that loads an Excel workbook, configures HtmlSaveOptions to include tooltip text, saves the result to a MemoryStream, and returns the UTF‑8 encoded HTML string while handling missing files and conversion errors.
    public static class AsposeCellsHelper
    {
        /// <param name="filePath">Full path to the Excel file to load.</param>
        /// <returns>HTML representation of the workbook with tooltips enabled.</returns>
        public static string LoadWorkbookAndGetHtml(string filePath)
        {
            // Ensure the file exists before attempting to load.
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");

            try
            {
                // Load the workbook from the given file path.
                Workbook workbook = new Workbook(filePath);

                // Configure HTML save options to include tooltip text.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    AddTooltipText = true
                };

                // Save the workbook to a memory stream using the configured options.
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, htmlOptions);
                    // Convert the memory stream contents to a UTF-8 encoded string.
                    return Encoding.UTF8.GetString(ms.ToArray());
                }
            }
            catch (Exception ex)
            {
                // Wrap any exception with a more descriptive message.
                throw new InvalidOperationException("Failed to load workbook or convert to HTML.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect a single argument: the path to the Excel file.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <excel-file-path>");
                return;
            }

            string filePath = args[0];

            try
            {
                string html = AsposeCellsHelper.LoadWorkbookAndGetHtml(filePath);
                Console.WriteLine(html);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
