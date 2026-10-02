// Title: Hide rows 5‑10 in an Excel worksheet and save as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to hide rows five through ten in the first worksheet of an Excel file and then export the workbook to a PDF document. | Provide a C# example that verifies an existing .xlsx file, sets the Row.IsHidden property for rows five to ten with Aspose.Cells, and saves the result as a PDF.
// Common Searches: Aspose.Cells C# hide specific rows before converting Excel to PDF | how to hide rows 5‑10 in an Excel file using Aspose.Cells and export to PDF | C# set Row.IsHidden for a range and save workbook as PDF with Aspose.Cells | convert Excel to PDF while rows are hidden using Aspose.Cells .NET
// Tags: Aspose.Cells hide rows Row.IsHidden | Aspose.Cells export hidden rows to PDF | C# hide Excel rows before PDF conversion | Aspose.Cells row visibility PDF output | Excel to PDF with hidden rows Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for input.xlsx, loads it with Aspose.Cells, hides rows 5‑10 (zero‑based indices 4‑9) by setting each Row.IsHidden to true on the first worksheet, and then saves the workbook as output.pdf using the PDF save format, with basic error handling.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide rows 5 through 10 (zero‑based index: 4 to 9)
            for (int rowIndex = 4; rowIndex <= 9; rowIndex++)
            {
                // Use the IsHidden property of the Row object
                sheet.Cells.Rows[rowIndex].IsHidden = true;
            }

            // Export the modified workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
