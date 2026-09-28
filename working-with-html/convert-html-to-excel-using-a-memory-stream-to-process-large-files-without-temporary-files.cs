// Title: Convert HTML to XLSX in C# using Aspose.Cells with only MemoryStream (no temporary files)
// AI Prompts: Write a C# method that receives a byte[] of HTML, loads it into an Aspose.Cells Workbook via a MemoryStream, and returns the workbook as an XLSX byte array. | Show how to read an HTML file into a byte array, convert it to an Excel workbook with Aspose.Cells, and write the resulting XLSX bytes to disk, using only in‑memory streams. | Provide sample code that includes error handling for converting large HTML content to Excel with Aspose.Cells without creating any temporary files, employing MemoryStream for both input and output.
// Common Searches: aspocells c# convert html string to xlsx using memory stream | how to load html into Aspose.Cells workbook from byte array | convert large html file to excel in C# without temporary files | save Aspose.Cells workbook to byte array instead of file | c# Aspose.Cells LoadOptions for html format memory stream
// Tags: Aspose.Cells load HTML from MemoryStream | Aspose.Cells save workbook as XLSX byte array | C# in‑memory HTML to Excel conversion | Aspose.Cells LoadOptions HTML format | MemoryStream based Excel generation with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates a ConvertHtmlToExcel method that loads HTML bytes into an Aspose.Cells Workbook via a MemoryStream, saves the workbook as an XLSX byte array, and shows a console example that reads an HTML file and writes the resulting Excel file—all without creating temporary files.
public class HtmlToExcelConverter
{
    /// <param name="htmlBytes">Byte array containing the HTML file.</param>
    /// <returns>Byte array of the generated Excel file (XLSX format).</returns>
    public byte[] ConvertHtmlToExcel(byte[] htmlBytes)
    {
        // Load the HTML from a memory stream without creating any temporary files.
        using (var inputStream = new MemoryStream(htmlBytes))
        {
            // Create a Workbook instance by loading the HTML stream.
            var loadOptions = new LoadOptions(LoadFormat.Html);
            var workbook = new Workbook(inputStream, loadOptions);

            // Save the workbook to another memory stream in XLSX format.
            using (var outputStream = new MemoryStream())
            {
                workbook.Save(outputStream, SaveFormat.Xlsx);
                // Return the Excel file as a byte array.
                return outputStream.ToArray();
            }
        }
    }
}

public class Program
{
    // Entry point for the console application.
    public static void Main(string[] args)
    {
        try
        {
            // Example usage: convert an HTML file to Excel.
            string htmlPath = "sample.html";
            string excelPath = "output.xlsx";

            // Ensure the input HTML file exists.
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Read HTML file into a byte array.
            byte[] htmlBytes = File.ReadAllBytes(htmlPath);

            // Perform conversion.
            var converter = new HtmlToExcelConverter();
            byte[] excelBytes = converter.ConvertHtmlToExcel(htmlBytes);

            // Write the resulting Excel file.
            File.WriteAllBytes(excelPath, excelBytes);
            Console.WriteLine($"Conversion successful. Excel file saved to: {excelPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
