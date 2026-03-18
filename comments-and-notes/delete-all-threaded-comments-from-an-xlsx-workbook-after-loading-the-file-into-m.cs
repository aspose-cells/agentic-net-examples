using System;
using Aspose.Cells;

namespace DeleteThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook into memory
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet and clear all comments (including threaded comments)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clears all comments in the worksheet
                sheet.ClearComments();
            }

            // Save the workbook after comments have been removed
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}