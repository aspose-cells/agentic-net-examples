using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            ImportXlsxWithMapping.Run();
        }
    }

    public class ImportXlsxWithMapping
    {
        public static void Run()
        {
            // Load the source XLSX file
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Create a new workbook for the destination
            Workbook destinationWorkbook = new Workbook();

            // Access the first worksheets of both workbooks
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Define the source range (A1:C3) and the destination start cell (E5)
            int srcStartRow = 0;
            int srcStartColumn = 0;
            int rowCount = 3;    // rows 0,1,2  (A1:C3)
            int columnCount = 3; // columns 0,1,2

            int destStartRow = 4;    // row 5 in Excel (E5)
            int destStartColumn = 4; // column 5 in Excel (E)

            // Copy each cell value from the source range to the destination range
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    // Get the source cell (always returns a Cell instance)
                    Cell srcCell = sourceSheet.Cells[srcStartRow + i, srcStartColumn + j];

                    // Get the destination cell (creates if it does not exist)
                    Cell destCell = destinationSheet.Cells[destStartRow + i, destStartColumn + j];

                    // Transfer the value (preserves data type)
                    destCell.PutValue(srcCell.Value);
                }
            }

            // Save the resulting workbook
            destinationWorkbook.Save("mapped_output.xlsx");
        }
    }
}