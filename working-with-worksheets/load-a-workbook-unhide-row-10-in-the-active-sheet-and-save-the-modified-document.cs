// Title: Unhide row 10 in the active worksheet of an Excel workbook using Aspose.Cells for .NET and save the result
// AI Prompts: Write C# code that opens an existing XLSX file with Aspose.Cells, makes row 10 visible in the currently active sheet, and writes the changes to a new file. | Show how to verify the presence of an input Excel file, access the active worksheet, set the IsHidden property of row 10 to false, and save the workbook using Aspose.Cells for .NET with exception handling. | Provide a step‑by‑step example that loads a workbook, retrieves the active worksheet, unhides row index 9, and saves the updated workbook to a specified output path.
// Common Searches: Aspose.Cells C# unhide specific row in active worksheet and save workbook | How to make row 10 visible in an Excel file using Aspose.Cells for .NET | C# code to load XLSX, change row hidden flag, and export to new file with Aspose.Cells | Unhide hidden rows programmatically with Aspose.Cells .NET API
// Tags: Aspose.Cells unhide row | C# Aspose.Cells modify row visibility | Excel workbook row IsHidden property .NET | active worksheet row manipulation Aspose.Cells | save modified XLSX with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads 'input.xlsx', unhides row 10 (zero‑based index 9) in the active worksheet, and saves the updated workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the active worksheet
            var worksheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Unhide row 10 (zero‑based index 9)
            var row = worksheet.Cells.Rows[9];
            row.IsHidden = false; // Correct property to control row visibility

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
