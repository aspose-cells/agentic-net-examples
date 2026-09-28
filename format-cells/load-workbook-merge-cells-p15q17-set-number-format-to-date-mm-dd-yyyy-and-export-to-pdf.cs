// Title: Merge cells P15:Q17, set a mm-dd-yyyy date format, and save the workbook as PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Combine cells P15 through Q17, assign a mm-dd-yyyy number format to the merged block, and generate a PDF file using Aspose.Cells in C#. | Create a style with a custom date pattern, apply it to a merged range, and export the Excel workbook to PDF via Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# merge specific cells and keep date formatting when converting to PDF | how to set custom number format for a merged range before PDF export using Aspose.Cells | C# example for applying mm-dd-yyyy style to merged cells and saving as PDF with Aspose.Cells | preserve date format on merged cells during Excel to PDF conversion in .NET | Aspose.Cells save workbook as PDF after formatting merged cells P15 Q17
// Tags: merge cell range Aspose.Cells C# | custom date number format mm-dd-yyyy Aspose.Cells | export workbook to PDF Aspose.Cells .NET | apply style to merged cells Aspose.Cells | Excel to PDF conversion preserving formatting Aspose.Cells | cell area styling Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// Loads input.xlsx, merges cells P15:Q17, applies a custom mm-dd-yyyy date style to the merged area, and saves the result as output.pdf using Aspose.Cells for .NET.
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells P15:Q17 (zero‑based indices)
            sheet.Cells.Merge(14, 15, 3, 2);

            // Create a style with the desired date format
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "mm-dd-yyyy";

            // Apply the date format only (preserve other formatting)
            StyleFlag flag = new StyleFlag { NumberFormat = true };

            // Define the merged range
            CellArea mergedArea = new CellArea
            {
                StartRow = 14,
                StartColumn = 15,
                EndRow = 16,
                EndColumn = 16
            };

            // Apply the style to each cell in the merged area
            for (int row = mergedArea.StartRow; row <= mergedArea.EndRow; row++)
            {
                for (int col = mergedArea.StartColumn; col <= mergedArea.EndColumn; col++)
                {
                    sheet.Cells[row, col].SetStyle(dateStyle, flag);
                }
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
