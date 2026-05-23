using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsUnionRangeDeleteDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data in two separate areas
                // First range: A1:B2
                cells["A1"].PutValue("A1");
                cells["A2"].PutValue("A2");
                cells["B1"].PutValue("B1");
                cells["B2"].PutValue("B2");

                // Second range: D4:E5
                cells["D4"].PutValue("D4");
                cells["D5"].PutValue("D5");
                cells["E4"].PutValue("E4");
                cells["E5"].PutValue("E5");

                // Create a UnionRange that represents both non‑contiguous ranges
                // The address string uses commas to separate individual ranges
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:B2,D4:E5", 0);

                // Delete each range in the union. The UnionRange defines all ranges to be removed in one logical step.
                foreach (AsposeRange r in unionRange.Ranges)
                {
                    // Delete the range and shift cells up to fill the gap
                    cells.DeleteRange(r.FirstRow, r.FirstColumn, r.RowCount, r.ColumnCount, ShiftType.Up);
                }

                // Define output file path
                string outputPath = "UnionRangeDeleteResult.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}