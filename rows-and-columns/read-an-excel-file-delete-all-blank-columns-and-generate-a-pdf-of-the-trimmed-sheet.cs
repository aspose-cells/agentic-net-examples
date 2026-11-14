// Title: C# – Remove all blank columns from an Excel worksheet and export the cleaned sheet to PDF using Aspose.Cells
// Common Searches: Aspose.Cells C# delete empty columns from worksheet before PDF export | How to trim blank columns in an .xlsx and save as PDF with Aspose.Cells .NET | C# code to remove columns that contain only null or whitespace values in Excel using Aspose.Cells | Delete columns with no data in an Excel workbook and generate PDF via Aspose.Cells | Export a trimmed Excel sheet to PDF without blank columns using Aspose.Cells for .NET
// Tags: blank column removal Aspose.Cells | worksheet column cleanup C# | PDF conversion after column deletion Aspose.Cells | detect empty column Aspose.Cells | save trimmed workbook as PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates over columns from the last used index backward, deletes any column that contains only null or whitespace values using a helper method, and then saves the resulting trimmed worksheet as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "trimmed_output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook does not contain any worksheets.");
                return;
            }

            // Work with the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last used column index
            int maxColumn = sheet.Cells.MaxDataColumn;

            // Iterate backwards to safely delete columns without affecting indices
            for (int col = maxColumn; col >= 0; col--)
            {
                if (IsColumnEmpty(sheet, col))
                {
                    // Delete the blank column
                    sheet.Cells.DeleteColumn(col);
                }
            }

            // Save the trimmed worksheet as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Successfully saved trimmed PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static bool IsColumnEmpty(Worksheet sheet, int columnIndex)
    {
        // Get the maximum used row index for the worksheet
        int maxRow = sheet.Cells.MaxDataRow;

        // Scan each cell in the column up to the last used row
        for (int row = 0; row <= maxRow; row++)
        {
            Cell cell = sheet.Cells[row, columnIndex];

            // Consider a cell empty if its value is null or an empty/whitespace string
            if (cell.Value != null)
            {
                if (cell.Value is string s)
                {
                    if (!string.IsNullOrWhiteSpace(s))
                        return false;
                }
                else
                {
                    return false;
                }
            }
        }

        // No data found in the column
        return true;
    }
}
