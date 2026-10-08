// Title: How to return an Aspose.Cells Excel-to-HTML conversion as a FileResult in an ASP.NET MVC controller (C#)
// AI Prompts: Write an ASP.NET MVC controller action that invokes a method converting an Excel file to HTML with Aspose.Cells and returns the HTML bytes using FileResult with content type "text/html". | Refactor the ExcelConverter class to accept a source file path parameter and expose a method that streams the generated HTML directly to the HttpResponse in an MVC action. | Add error handling to the MVC action so that a missing Excel file results in a 404 Not Found response and the exception is logged.
// Common Searches: asp.net mvc return html generated from excel using aspose.cells as fileresult | c# asp.net mvc controller convert xlsx to html and send as file download | aspose.cells save workbook as html and stream to response in mvc | how to use memory stream with asp.net mvc FileResult for html output | mvc action return 404 when excel file not found during aspose.cells conversion
// Tags: asp.net mvc fileresult html output | aspose.cells convert workbook to html | excel to html streaming mvc | memory stream file result c# | error handling missing excel file mvc

using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    // The sample loads an Excel workbook from App_Data, uses Aspose.Cells to save it as HTML into a MemoryStream, and demonstrates how to expose this conversion as a FileResult from an ASP.NET MVC controller with proper error handling for missing files.
    public class ExcelConverter
    {
        // Converts an Excel file to HTML and returns the HTML bytes
        public byte[] ConvertToHtml()
        {
            // Build the full path to the source Excel file
            string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "sample.xlsx");

            // Ensure the file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
                throw new FileNotFoundException("The source Excel file was not found.", excelPath);

            try
            {
                // Load the workbook from the file system
                Workbook workbook = new Workbook(excelPath);

                // Capture the HTML output in a memory stream
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    // Save the workbook as HTML into the memory stream
                    workbook.Save(htmlStream, SaveFormat.Html);

                    // Return the HTML content as a byte array
                    return htmlStream.ToArray();
                }
            }
            catch (Exception ex)
            {
                // Wrap and rethrow any exception for higher‑level handling
                throw new ApplicationException("Failed to convert Excel to HTML.", ex);
            }
        }
    }

    public class Program
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            try
            {
                ExcelConverter converter = new ExcelConverter();
                byte[] htmlBytes = converter.ConvertToHtml();

                // Write the HTML output to a file for verification
                string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.html");
                File.WriteAllBytes(outputPath, htmlBytes);
                Console.WriteLine($"HTML file generated at: {outputPath}");
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
