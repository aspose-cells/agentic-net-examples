using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveWorksheetByNameDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add worksheets with distinct names
            wb.Worksheets.Add("SheetA");
            wb.Worksheets.Add("SheetB");
            wb.Worksheets.Add("SheetC");

            // Show worksheet count before removal
            Console.WriteLine($"Worksheets before removal: {wb.Worksheets.Count}");

            // Remove the worksheet named "SheetB"
            Worksheet sheetToRemove = wb.Worksheets["SheetB"];
            if (sheetToRemove != null)
            {
                wb.Worksheets.RemoveAt(sheetToRemove.Index);
            }

            // Show worksheet count after removal
            Console.WriteLine($"Worksheets after removal: {wb.Worksheets.Count}");

            // Save the workbook to a file
            wb.Save("RemovedSheetDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveWorksheetByNameDemo.Run();
        }
    }
}