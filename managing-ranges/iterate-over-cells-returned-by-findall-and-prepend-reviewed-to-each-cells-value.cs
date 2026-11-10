// Title: How to prepend "Reviewed:" to every non‑empty cell in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells prepend text to each non‑empty cell in a worksheet | How to add a prefix to all populated Excel cells with Aspose.Cells .NET | Iterate over used range and modify cell values using Aspose.Cells API | Bulk update cell strings in an Excel file with Aspose.Cells C# | Prefix cell content with "Reviewed:" using Aspose.Cells library
// Tags: prepend prefix to cell values Aspose.Cells | bulk modify non‑empty Excel cells C# | iterate used range worksheet Aspose.Cells | update cell string content .NET Excel | add text to Excel cell values Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates a new one), accesses the first worksheet, scans the used range for cells that contain non‑empty string values, prepends the literal "Reviewed:" to each cell's text, and saves the modified workbook. It demonstrates safe iteration over populated cells and bulk text updates with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range.
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;

            // Iterate through all cells that contain a value.
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = worksheet.Cells[row, col];
                    if (cell != null && cell.Value != null && !string.IsNullOrEmpty(cell.StringValue))
                    {
                        string originalValue = cell.StringValue;
                        cell.PutValue("Reviewed:" + originalValue);
                    }
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
