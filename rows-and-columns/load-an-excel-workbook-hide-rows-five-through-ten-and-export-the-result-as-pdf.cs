// Title: Hide rows 5‑10 in an Excel worksheet and save as PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide rows 5 to 10 before exporting workbook to PDF | How to programmatically hide multiple rows in an Excel file with Aspose.Cells and generate a PDF | C# example for setting Row.IsHidden property and saving the workbook as PDF using Aspose.Cells | Convert Excel to PDF after hiding specific rows with Aspose.Cells .NET | Aspose.Cells hide rows range and export to PDF in a C# console application
// Tags: row.IsHidden Aspose.Cells C# | hide multiple rows Excel .NET | Excel to PDF conversion after row visibility change | create output directory before saving PDF Aspose.Cells | load workbook and modify row visibility Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program checks for the presence of 'input.xlsx', loads it with Aspose.Cells, accesses the first worksheet, hides rows 5 through 10 by setting each Row's IsHidden property to true, ensures the output folder exists, and then saves the modified workbook as 'output.pdf' in PDF format, handling errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide rows 5 through 10 (zero‑based indices 4 to 9)
            for (int rowIndex = 4; rowIndex <= 9; rowIndex++)
            {
                // Use the IsHidden property of Row objects
                sheet.Cells.Rows[rowIndex].IsHidden = true;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the modified workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
