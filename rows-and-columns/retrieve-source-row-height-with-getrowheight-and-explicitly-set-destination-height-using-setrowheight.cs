using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRowHeightExample
{
    public class RowHeightTransfer
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (source) and set a custom row height
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Set height for row index 2 (third row)
            sourceCells.SetRowHeight(2, 35.0);
            double sourceRowHeight = sourceCells.GetRowHeight(2);
            Console.WriteLine($"Source row (2) height: {sourceRowHeight}");

            // Create a new workbook (destination) and apply the source row height
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            Cells destinationCells = destinationSheet.Cells;

            destinationCells.SetRowHeight(5, sourceRowHeight);
            Console.WriteLine($"Destination row (5) height set to: {destinationCells.GetRowHeight(5)}");

            // Save workbooks – ensure the directory is writable
            string sourcePath = "SourceWorkbook.xlsx";
            string destPath = "DestinationWorkbook.xlsx";

            sourceWorkbook.Save(sourcePath);
            destinationWorkbook.Save(destPath);

            Console.WriteLine($"Workbooks saved: {Path.GetFullPath(sourcePath)}, {Path.GetFullPath(destPath)}");
        }
    }
}