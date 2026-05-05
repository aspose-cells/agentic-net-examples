using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRemoveWorksheetDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet named "Sheet1")
            Workbook workbook = new Workbook();

            // Add additional worksheets for demonstration
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Display worksheet count before removal
            Console.WriteLine($"Worksheets before removal: {workbook.Worksheets.Count}");

            // Remove the worksheet with the specified name
            // The RemoveAt(string) method deletes the sheet identified by its name.
            string sheetToRemove = "Sheet2";
            workbook.Worksheets.RemoveAt(sheetToRemove);

            // Display worksheet count after removal
            Console.WriteLine($"Worksheets after removal of \"{sheetToRemove}\": {workbook.Worksheets.Count}");

            // Save the modified workbook to a file
            string outputPath = "RemovedSheetDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}