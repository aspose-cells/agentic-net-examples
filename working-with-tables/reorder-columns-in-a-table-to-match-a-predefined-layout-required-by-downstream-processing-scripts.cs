using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Needed for ListObject

namespace AsposeCellsColumnReorderDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a workbook and populate it with sample data (including header row)
                Workbook workbook = new Workbook();
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "Source";

                // Headers: ColA, ColB, ColC, ColD
                srcSheet.Cells["A1"].PutValue("ColA");
                srcSheet.Cells["B1"].PutValue("ColB");
                srcSheet.Cells["C1"].PutValue("ColC");
                srcSheet.Cells["D1"].PutValue("ColD");

                // Sample rows
                for (int row = 2; row <= 5; row++)
                {
                    srcSheet.Cells[row - 1, 0].PutValue($"A{row - 1}");
                    srcSheet.Cells[row - 1, 1].PutValue($"B{row - 1}");
                    srcSheet.Cells[row - 1, 2].PutValue($"C{row - 1}");
                    srcSheet.Cells[row - 1, 3].PutValue($"D{row - 1}");
                }

                // Desired column order (0‑based indexes of the source columns)
                // Example: reorder to C, A, D, B
                int[] desiredOrder = new int[] { 2, 0, 3, 1 };

                // Add a new worksheet that will hold the reordered columns
                Worksheet destSheet = workbook.Worksheets.Add("Reordered");

                // Copy each column from the source sheet to the destination sheet according to the desired order
                for (int destCol = 0; destCol < desiredOrder.Length; destCol++)
                {
                    int srcCol = desiredOrder[destCol];
                    destSheet.Cells.CopyColumns(srcSheet.Cells, srcCol, destCol, 1);
                }

                // Create a table (ListObject) on the destination sheet covering the data range
                int totalRows = srcSheet.Cells.MaxDataRow + 1; // include header row
                int totalCols = desiredOrder.Length;
                int tableIndex = destSheet.ListObjects.Add(0, 0, totalRows - 1, totalCols - 1, true);
                ListObject destTable = destSheet.ListObjects[tableIndex];
                destTable.DisplayName = "ReorderedTable";

                // Ensure the ListObject column names reflect the header cells after reordering
                destTable.UpdateColumnName();

                // Save the workbook with the reordered columns
                string outputPath = "ReorderedColumns.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}