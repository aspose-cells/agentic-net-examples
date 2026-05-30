using System;
using System.IO;
using Aspose.Cells;

class HtmlToExcelConverter
{
    static void Main()
    {
        // Paths for the source HTML file and the destination Excel file
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        try
        {
            // Verify that the HTML source file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"HTML source file not found: {htmlPath}");

            // Configure HTML load options (keep numeric precision)
            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions
            {
                KeepPrecision = true
            };

            // Load the HTML file into a workbook using the HtmlLoadOptions
            Workbook workbook = new Workbook(htmlPath, htmlLoadOptions);

            // Convert string values that can be interpreted as numbers to numeric types
            workbook.Worksheets[0].Cells.ConvertStringToNumericValue();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook as an Excel file (XLSX format)
            workbook.Save(excelPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Output any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}