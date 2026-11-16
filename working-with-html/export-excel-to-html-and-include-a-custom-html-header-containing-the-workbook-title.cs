// Title: Export Excel workbook to HTML and insert a custom <h1> header with the workbook title using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export Excel to HTML with document property title | How to add an H1 heading after the <body> tag when saving Excel as HTML with Aspose.Cells | Insert workbook title into generated HTML from an .xlsx file using Aspose.Cells | Aspose.Cells HtmlSaveOptions save to memory stream then modify HTML output
// Tags: Aspose.Cells HtmlSaveOptions add heading | Excel to HTML export document property title | C# insert H1 after body tag | MemoryStream HTML generation Aspose.Cells | Read workbook BuiltInDocumentProperties Title C#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an .xlsx file with Aspose.Cells, reads the workbook's Title built‑in property (defaulting to "Workbook" if missing), saves the workbook to a MemoryStream as HTML using HtmlSaveOptions, injects an <h1> element containing the title immediately after the <body> tag, and writes the modified HTML to an output file, with comprehensive error handling at each step.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.html";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found.");
                    return;
                }

                // Load the Excel workbook
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputFile);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Retrieve the title from built‑in document properties
                string title = workbook.BuiltInDocumentProperties.Title;
                if (string.IsNullOrEmpty(title))
                {
                    title = "Workbook";
                }

                // Configure HTML export options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Save the workbook to a memory stream first
                byte[] htmlBytes;
                try
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        workbook.Save(ms, htmlOptions);
                        htmlBytes = ms.ToArray();
                    }
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save HTML to memory: {saveEx.Message}");
                    return;
                }

                // Insert a custom header (title) into the generated HTML
                string htmlContent = Encoding.UTF8.GetString(htmlBytes);
                int bodyPos = htmlContent.IndexOf("<body>", StringComparison.OrdinalIgnoreCase);
                if (bodyPos != -1)
                {
                    int insertPos = bodyPos + "<body>".Length;
                    htmlContent = htmlContent.Insert(insertPos, $"<h1>{title}</h1>");
                }

                // Write the final HTML to the output file
                try
                {
                    File.WriteAllText(outputFile, htmlContent, Encoding.UTF8);
                    Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
                }
                catch (Exception writeEx)
                {
                    Console.WriteLine($"Failed to write HTML file: {writeEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
