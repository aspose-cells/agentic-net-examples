using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
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

                // Define the original range (e.g., A1:C3)
                Aspose.Cells.Range originalRange = cells.CreateRange("A1:C3");

                // Add comments to some cells inside the original range
                int idxA1 = worksheet.Comments.Add("A1");
                worksheet.Comments[idxA1].Note = "Comment on A1";

                int idxB2 = worksheet.Comments.Add("B2");
                worksheet.Comments[idxB2].Note = "Comment on B2";

                int idxC3 = worksheet.Comments.Add("C3");
                worksheet.Comments[idxC3].Note = "Comment on C3";

                // Offset the range by 3 rows and 5 columns
                Aspose.Cells.Range offsetRange = originalRange.GetOffset(3, 5);

                // Add a comment to a cell inside the offset range to demonstrate clearing
                string topLeftCell = CellsHelper.CellIndexToName(offsetRange.FirstRow, offsetRange.FirstColumn);
                int idxOffset = worksheet.Comments.Add(topLeftCell);
                worksheet.Comments[idxOffset].Note = "Comment in offset range";

                // Clear all comments within the offset range
                offsetRange.ClearComments();

                // Save the workbook
                string outputPath = "OffsetClearCommentsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}