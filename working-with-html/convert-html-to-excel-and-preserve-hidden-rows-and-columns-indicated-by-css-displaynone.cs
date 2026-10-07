// Title: Convert HTML with CSS display:none rows and columns to an Excel workbook while preserving hidden elements using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an HTML file with Aspose.Cells, retains rows and columns hidden by CSS display:none, and saves the result as an XLSX workbook. | Show how to configure HtmlLoadOptions in Aspose.Cells to keep CSS‑based hidden rows and columns during HTML‑to‑Excel conversion. | Write a robust C# snippet that verifies the input HTML file, converts it to Excel while preserving hidden elements, and includes proper exception handling.
// Common Searches: Aspose.Cells C# preserve CSS display:none rows when converting HTML to XLSX | How to keep hidden columns from an HTML table in Excel output using Aspose.Cells | HtmlLoadOptions hide rows and columns based on CSS in .NET conversion | Convert HTML table with display:none cells to Excel while maintaining hidden state | C# Aspose.Cells convert HTML to Excel respecting CSS visibility rules
// Tags: Aspose.Cells hidden rows conversion | HtmlLoadOptions display:none handling | C# preserve hidden columns in Excel export | HTML table to XLSX visibility retention | Aspose.Cells workbook import CSS visibility

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an HTML file with Aspose.Cells using HtmlLoadOptions, creates a Workbook, and saves it as an XLSX file while preserving rows and columns hidden via CSS display:none, with file existence checks and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output paths
            string inputPath = @"C:\Path\To\Input.html";
            string outputPath = @"C:\Path\To\Output.xlsx";

            try
            {
                // Verify that the input HTML file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Set HTML load options (default options are used here)
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();

                // Load the HTML file into a new workbook using the specified options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Save the workbook as an Excel file (XLSX format)
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook successfully saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
