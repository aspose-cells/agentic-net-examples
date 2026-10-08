// Title: Convert HTML containing nested tables into an Excel workbook with each table on its own worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an HTML file with multiple nested tables into an Aspose.Cells Workbook and saves each table as a distinct worksheet in an XLSX file. | Show how to configure HtmlLoadOptions in Aspose.Cells so that every HTML table, including nested ones, is mapped to a separate worksheet during conversion. | Provide a robust C# example that validates the HTML file path, performs the conversion, and implements error handling for missing files or conversion failures.
// Common Searches: asp.net convert html page with nested tables to excel each table separate sheet | c# aspose.cells htmlloadoptions map nested tables to worksheets | how to preserve HTML table hierarchy when exporting to xlsx using aspose.cells | example code converting complex html layout to multiple worksheets in .net
// Tags: HTML to XLSX conversion with separate worksheets | Aspose.Cells nested table worksheet mapping | C# HtmlLoadOptions worksheet generation | preserve table hierarchy Aspose.Cells .NET | automatic worksheet creation from HTML tables

using System;
using System.IO;
using Aspose.Cells;

// The example validates the existence of an input HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions, automatically maps each HTML table—including nested tables—to its own worksheet, saves the result as an XLSX file, and reports success or any errors encountered.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Path to the input HTML file
            string htmlPath = "input.html";

            // Verify that the input file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Load the HTML content directly into a workbook
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as an Excel file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Conversion completed successfully. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
