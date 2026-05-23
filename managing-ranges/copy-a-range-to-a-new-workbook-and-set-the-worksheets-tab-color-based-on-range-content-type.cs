using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create source workbook and populate sample data
                Workbook sourceWorkbook = new Workbook(); // empty workbook
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                sourceSheet.Cells["A1"].PutValue("Text");
                sourceSheet.Cells["A2"].PutValue(123);
                sourceSheet.Cells["B1"].PutValue(DateTime.Now);
                sourceSheet.Cells["B2"].PutValue(456.78);

                // Define the source range to be copied
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:B2");

                // Create destination workbook
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Define the destination range (same size as source)
                AsposeRange destRange = destSheet.Cells.CreateRange("A1:B2");

                // Copy the source range (data, formulas, formatting, etc.) to the destination range
                destRange.Copy(sourceRange);

                // Determine content types present in the source range
                bool hasNumeric = false;
                bool hasString = false;
                bool hasDate = false;

                for (int row = 0; row < sourceRange.RowCount; row++)
                {
                    for (int col = 0; col < sourceRange.ColumnCount; col++)
                    {
                        Cell cell = sourceRange[row, col];
                        if (cell.Type == CellValueType.IsNumeric)
                            hasNumeric = true;
                        else if (cell.Type == CellValueType.IsString)
                            hasString = true;
                        else if (cell.Type == CellValueType.IsDateTime)
                            hasDate = true;
                    }
                }

                // Set worksheet tab color based on detected content
                if (hasNumeric && !hasString && !hasDate)
                    destSheet.TabColor = Color.Green;          // only numeric values
                else if (hasString && !hasNumeric && !hasDate)
                    destSheet.TabColor = Color.Blue;           // only text values
                else if (hasDate && !hasNumeric && !hasString)
                    destSheet.TabColor = Color.Orange;         // only dates
                else
                    destSheet.TabColor = Color.Gray;           // mixed or other types

                // Save the destination workbook
                string outputPath = "CopiedRangeWithTabColor.xlsx";
                destWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}