// Title: Copy a column from one worksheet to another in an Excel file using Aspose.Cells for .NET while preserving column width, data types, styles, and formulas
// AI Prompts: Use Aspose.Cells in C# to copy column C from Sheet1 to column F in Sheet2, keeping the original column width, cell values, formatting, and any formulas. | Write C# code with Aspose.Cells that transfers a source column to a target column on a different worksheet, ensuring data types, styles, and column width are preserved.
// Common Searches: aspnet copy column between worksheets preserving column width Aspose.Cells | c# Aspose.Cells copy column with formulas and formatting to another sheet | how to keep data types when moving Excel column using Aspose.Cells .NET | copy column from Sheet1 to Sheet2 while retaining column width in Aspose.Cells
// Tags: copy column across worksheets Aspose.Cells | preserve column width Aspose.Cells C# | retain cell formulas Aspose.Cells | transfer column formatting Excel .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads (or creates) an input.xlsx workbook, ensures Sheet1 and Sheet2 exist, then copies column C (index 2) from Sheet1 to column F (index 5) in Sheet2. It preserves the source column's width, each cell's value, style, and any formulas before saving the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a minimal workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets.Clear();
                tempWb.Worksheets.Add("Sheet1");
                tempWb.Worksheets.Add("Sheet2");
                tempWb.Save(inputPath);
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Retrieve source and destination worksheets; create them if missing.
            Worksheet srcSheet = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];
            Worksheet destSheet = workbook.Worksheets["Sheet2"] ?? workbook.Worksheets[1];

            // Column indexes (0‑based). Example: copy column C (index 2) to column F (index 5)
            int srcColumnIndex = 2;
            int destColumnIndex = 5;

            // Preserve column width
            double srcColumnWidth = srcSheet.Cells.GetColumnWidth(srcColumnIndex);
            destSheet.Cells.SetColumnWidth(destColumnIndex, srcColumnWidth);

            // Determine the last row that contains data in the source column
            int maxRow = srcSheet.Cells.MaxDataRow;

            // Copy cells while preserving data types, values, styles, and formulas
            for (int row = 0; row <= maxRow; row++)
            {
                Cell srcCell = srcSheet.Cells[row, srcColumnIndex];
                Cell destCell = destSheet.Cells[row, destColumnIndex];

                // Copy the cell value (preserves the underlying data type)
                destCell.PutValue(srcCell.Value);

                // Copy the cell style (includes number format, font, alignment, etc.)
                destCell.SetStyle(srcCell.GetStyle());

                // Copy formula if present
                if (!string.IsNullOrEmpty(srcCell.Formula))
                {
                    destCell.Formula = srcCell.Formula;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
