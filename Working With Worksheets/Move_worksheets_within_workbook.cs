using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            MoveWorksheetsDemo.Run();
        }
    }

    public class MoveWorksheetsDemo
    {
        public static void Run()
        {
            // Create a new workbook and remove the default sheet
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            // Add three worksheets with custom names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Add a new worksheet that will be moved
            Worksheet movedSheet = workbook.Worksheets.Add("MovedSheet");

            // Move the newly added sheet to index 1 (second position)
            movedSheet.MoveTo(1);

            // Swap Sheet1 (index 0) with Sheet3 (index 3)
            workbook.Worksheets.SwapSheet(0, 3);

            // Save the workbook to verify the changes
            workbook.Save("MovedWorksheetsDemo.xlsx");
        }
    }
}