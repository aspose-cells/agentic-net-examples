using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

public class HtmlToExcelConverter
{
    /// <summary>
    /// Converts an HTML file to an Excel workbook and saves it to the specified output path.
    /// </summary>
    /// <param name="htmlFilePath">Full path to the source HTML file.</param>
    /// <param name="outputExcelPath">Full path where the resulting XLSX file will be saved.</param>
    public static void ConvertHtmlToExcel(string htmlFilePath, string outputExcelPath)
    {
        try
        {
            // Verify that the source HTML file exists to avoid FileNotFoundException.
            if (!File.Exists(htmlFilePath))
                throw new FileNotFoundException($"HTML file not found: {htmlFilePath}");

            // Load the HTML workbook (Aspose.Cells can load HTML directly).
            var workbook = new Workbook(htmlFilePath);

            // Ensure the output directory exists.
            var outputDir = Path.GetDirectoryName(outputExcelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook in XLSX format.
            workbook.Save(outputExcelPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Log or rethrow as needed; for this example we simply write to console.
            Console.Error.WriteLine($"Error during conversion: {ex.Message}");
            throw;
        }
    }
}

// Example usage
class Program
{
    static void Main()
    {
        // Path to the source HTML file (local file system).
        string htmlPath = "sample.html";

        // Desired output Excel file path.
        string outputPath = "converted.xlsx";

        try
        {
            HtmlToExcelConverter.ConvertHtmlToExcel(htmlPath, outputPath);
            Console.WriteLine("HTML has been successfully converted to Excel.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Conversion failed: {ex.Message}");
        }
    }
}