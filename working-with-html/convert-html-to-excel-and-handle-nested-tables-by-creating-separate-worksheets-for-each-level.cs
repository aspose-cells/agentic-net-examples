// Title: Convert an HTML file with tables to an Excel workbook using Aspose.Cells in C#
// Common Searches: asp.net c# load html string into workbook with Aspose.Cells HtmlLoadOptions | how to read an HTML file into a MemoryStream for Aspose.Cells conversion | save Aspose.Cells workbook to a specific directory in C# | check if HTML file exists before converting to Excel with Aspose.Cells | create missing output folder when saving an Aspose.Cells workbook
// Tags: Aspose.Cells HtmlLoadOptions memory stream | Aspose.Cells convert html tables to excel workbook | Aspose.Cells workbook save as xlsx | C# verify html file existence before conversion | C# ensure output directory for workbook save

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example reads an HTML file, loads its content into an Aspose.Cells Workbook via a UTF‑8 MemoryStream and HtmlLoadOptions, ensures the target folder exists, and saves the workbook as an .xlsx file, with basic error handling.
class HtmlToExcelConverter
{
    static void Main(string[] args)
    {
        // Define input HTML and output Excel paths
        string inputHtmlPath = @"C:\Temp\input.html";
        string outputExcelPath = @"C:\Temp\output.xlsx";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(inputHtmlPath))
                throw new FileNotFoundException("Input HTML file not found.", inputHtmlPath);

            // Read the entire HTML content
            string htmlContent = File.ReadAllText(inputHtmlPath);

            // Load HTML content into a workbook using a memory stream and HtmlLoadOptions
            Workbook workbook;
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(ms, loadOptions);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputExcelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to the specified file
            workbook.Save(outputExcelPath);
            Console.WriteLine($"Excel file saved successfully to: {outputExcelPath}");
        }
        catch (Exception ex)
        {
            // Log the error details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
