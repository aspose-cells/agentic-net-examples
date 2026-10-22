// Title: Convert HTML to XLSX in C# using Aspose.Cells with only MemoryStream (no temporary files)
// AI Prompts: Create a C# method that accepts an HTML Stream, loads it into an Aspose.Cells Workbook, and returns a MemoryStream with the workbook saved as XLSX. | Show how to perform an HTML‑to‑Excel conversion entirely in memory with Aspose.Cells, avoiding any file system writes. | Provide example code that reads HTML from a file or string, generates an XLSX workbook via Aspose.Cells, and streams the result to another MemoryStream for further processing.
// Common Searches: aspnet core convert html stream to xlsx using aspose.cells without temporary files | c# load html into workbook from memory stream and save as excel in memory | how to process large html files to excel with aspose.cells using streams | convert html table to excel workbook in c# using only memory streams
// Tags: Aspose.Cells load HTML stream | Aspose.Cells save workbook to MemoryStream | HTML to XLSX conversion in memory | C# in‑memory HTML to Excel | stream‑based large HTML to Excel processing

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelApp
{
    // The example defines a HtmlToExcelConverter with a ConvertHtmlToExcel method that takes a Stream containing HTML, loads it into an Aspose.Cells Workbook via LoadOptions(LoadFormat.Html), saves the workbook to a MemoryStream in XLSX format, and returns that stream. The program demonstrates reading HTML from a file path or a string, converting it, and writing the resulting Excel file to disk, all while avoiding temporary files.
    public class HtmlToExcelConverter
    {
        /// <param name="htmlStream">Stream containing the HTML data.</param>
        /// <returns>MemoryStream containing the generated Excel file (XLSX format).</returns>
        public MemoryStream ConvertHtmlToExcel(Stream htmlStream)
        {
            if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));

            // Ensure the input stream is at the beginning
            if (htmlStream.CanSeek)
                htmlStream.Position = 0;

            // Load HTML into a workbook
            Workbook workbook = new Workbook(htmlStream, new LoadOptions(LoadFormat.Html));

            // Save workbook to a memory stream in XLSX format
            MemoryStream excelStream = new MemoryStream();
            workbook.Save(excelStream, SaveFormat.Xlsx);
            excelStream.Position = 0;
            return excelStream;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Stream htmlInputStream;

                if (args.Length > 0 && File.Exists(args[0]))
                {
                    // Load HTML from file if path is provided and exists
                    htmlInputStream = new FileStream(args[0], FileMode.Open, FileAccess.Read);
                }
                else
                {
                    // Use sample HTML content
                    string htmlContent = "<html><body><table><tr><td>Data</td></tr></table></body></html>";
                    htmlInputStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent));
                }

                using (htmlInputStream)
                {
                    HtmlToExcelConverter converter = new HtmlToExcelConverter();
                    using (MemoryStream excelOutput = converter.ConvertHtmlToExcel(htmlInputStream))
                    {
                        // Write the Excel file to disk
                        string outputPath = "output.xlsx";
                        using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                        {
                            excelOutput.CopyTo(file);
                        }
                        Console.WriteLine($"Excel file saved to {Path.GetFullPath(outputPath)}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
