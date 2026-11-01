// Title: How to delete all blank columns from an Excel worksheet and export the cleaned sheet to PDF using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx workbook with Aspose.Cells, scan columns from the last to the first, delete any column whose cells are all Null, and then generate a PDF file. | In C#, use Aspose.Cells to find completely empty columns in the first worksheet, remove them safely, and export the resulting sheet as a PDF.
// Common Searches: C# Aspose.Cells remove columns that have only empty cells before PDF conversion | How to clean an Excel sheet by deleting empty columns using Aspose.Cells .NET | Export a trimmed worksheet to PDF after removing blank columns with Aspose.Cells | Iterate Excel columns backwards to safely delete empty columns in C# Aspose.Cells
// Tags: delete empty columns Aspose.Cells C# | export worksheet to PDF Aspose.Cells | reverse column iteration Aspose.Cells | blank column cleanup Excel .NET | save workbook as PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads "input.xlsx", scans each column from right to left, deletes columns that contain only null cells, and saves the resulting worksheet as "output.pdf" using Aspose.Cells.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust the index if needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the maximum used row to limit the scan
            int maxRow = cells.MaxDataRow;

            // Iterate columns from right to left to safely delete blank columns
            for (int col = cells.MaxDataColumn; col >= 0; col--)
            {
                bool isBlankColumn = true;

                // Check each cell in the column up to the last used row
                for (int row = 0; row <= maxRow; row++)
                {
                    // A cell is considered blank if its type is Null
                    if (cells[row, col].Type != CellValueType.IsNull)
                    {
                        isBlankColumn = false;
                        break;
                    }
                }

                // Delete the column if it is completely blank
                if (isBlankColumn)
                {
                    cells.DeleteColumn(col);
                }
            }

            // Save the trimmed worksheet as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Successfully saved PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
