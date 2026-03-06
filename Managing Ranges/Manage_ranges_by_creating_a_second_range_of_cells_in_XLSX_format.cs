using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class ManageRangesDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create the first range (A1:C3) and fill it with sample data
            AsposeRange sourceRange = cells.CreateRange("A1", "C3");
            for (int i = 0; i < sourceRange.RowCount; i++)
            {
                for (int j = 0; j < sourceRange.ColumnCount; j++)
                {
                    sourceRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Create a second range with the same dimensions starting at column E (index 4)
            AsposeRange destRange = cells.CreateRange(0, 4, sourceRange.RowCount, sourceRange.ColumnCount);

            // Copy only the cell values from the source range to the destination range
            destRange.CopyValue(sourceRange);

            // Assign names to the ranges (optional, useful for formulas)
            sourceRange.Name = "SourceRange";
            destRange.Name = "DestRange";

            // Save the workbook in XLSX format
            workbook.Save("ManagedRanges.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ManageRangesDemo.Run();
        }
    }
}