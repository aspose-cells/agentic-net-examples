using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            RemoveWorksheetsByIndex.Run();
        }
    }

    public class RemoveWorksheetsByIndex
    {
        public static void Run()
        {
            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Rename the default worksheet and add two more worksheets
            WorksheetCollection sheets = workbook.Worksheets;
            sheets[0].Name = "Sheet1";
            sheets.Add("Sheet2");
            sheets.Add("Sheet3");

            // Show count before removal
            Console.WriteLine($"Worksheets before removal: {sheets.Count}");

            // Remove the worksheet at index 1 (second worksheet)
            sheets.RemoveAt(1);

            // Show count after removal
            Console.WriteLine($"Worksheets after removal: {sheets.Count}");

            // Save the workbook
            workbook.Save("RemovedSheetDemo.xlsx");
        }
    }
}