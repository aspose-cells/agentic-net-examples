// Title: Convert string‑based numeric values to numbers in the first worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert numeric strings to numbers only on the first sheet | How to change cell type from string to double in worksheet 1 using Aspose.Cells | Replace string numbers with numeric values in an Excel workbook via Aspose.Cells .NET | Iterate used range and parse numeric strings in the first worksheet with Aspose.Cells C# | Save workbook after converting string numbers to numeric in the first worksheet Aspose.Cells
// Tags: string to numeric cell conversion Aspose.Cells | first worksheet processing Aspose.Cells .NET | iterate used range Excel Aspose.Cells | cell value type update C# Aspose.Cells | save workbook after value conversion Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file using Aspose.Cells, accesses the first worksheet, iterates through its used range, and converts any cell containing a numeric string into a true numeric value before saving the workbook. It also includes file‑existence checking and exception handling for Aspose.Cells and general errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook (no password). If the file is password‑protected,
            // a CellsException will be thrown and caught below.
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Iterate through each used cell and convert numeric strings to numbers
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    var cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue.Trim();
                        if (double.TryParse(text, out double numericValue))
                        {
                            cell.PutValue(numericValue);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
