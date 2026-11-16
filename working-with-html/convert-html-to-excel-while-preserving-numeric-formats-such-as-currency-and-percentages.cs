// Title: Convert HTML tables to Excel (.xlsx) while preserving currency and percentage formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load HTML with numeric data conversion to keep currency formatting | How to retain percentage values when converting HTML to XLSX with Aspose.Cells | Convert HTML file to Excel workbook preserving numeric formats using HtmlLoadOptions | C# example for converting HTML tables to .xlsx while keeping numeric types | Aspose.Cells ConvertNumericData true usage for HTML to Excel conversion
// Tags: HtmlLoadOptions ConvertNumericData true | MemoryStream HTML import Workbook Aspose.Cells | numeric data conversion HTML to Excel C# | currency percentage formatting preservation Aspose.Cells | HTML table to XLSX conversion using Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample reads an HTML file, enables HtmlLoadOptions.ConvertNumericData to treat numeric strings as numbers, loads the content into a Workbook via a MemoryStream, and saves it as an XLSX file, ensuring that currency and percentage values retain their numeric formatting in the resulting Excel workbook.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // Read the HTML content
            string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);

            // Configure HTML load options to convert numeric strings to numeric cells
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                ConvertNumericData = true
            };

            // Load the HTML content into a new workbook using a memory stream
            Workbook workbook;
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                workbook = new Workbook(ms, loadOptions);
            }

            // Path for the output Excel file
            string excelPath = "output.xlsx";

            // Save the workbook as an Excel file
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML has been converted to Excel and saved as '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
