using System;
using Aspose.Cells;

namespace AsposeCellsActiveSheetDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Add a second worksheet to demonstrate switching active sheets
            Worksheet secondSheet = workbook.Worksheets.Add("SecondSheet");

            // Set the first worksheet as the active sheet by index
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Access the active worksheet
            Worksheet activeWorksheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Set the active cell in the active worksheet to B2 (Worksheet.ActiveCell property)
            activeWorksheet.ActiveCell = "B2";

            // Optionally mark the worksheet as selected when the file is opened
            activeWorksheet.IsSelected = true;

            // Save the workbook (lifecycle save rule)
            string filePath = "ActiveSheetDemo.xlsx";
            workbook.Save(filePath);

            // Load the saved workbook to verify the active sheet and cell (lifecycle load rule)
            Workbook loadedWorkbook = new Workbook(filePath);

            // Retrieve the active sheet index and active cell after loading
            int activeIndex = loadedWorkbook.Worksheets.ActiveSheetIndex;
            string activeCell = loadedWorkbook.Worksheets[activeIndex].ActiveCell;

            Console.WriteLine($"Active Sheet Index after load: {activeIndex}");
            Console.WriteLine($"Active Cell in active sheet after load: {activeCell}");

            // Switch active sheet to the second worksheet and set a different active cell
            loadedWorkbook.Worksheets.ActiveSheetIndex = loadedWorkbook.Worksheets["SecondSheet"].Index;
            loadedWorkbook.Worksheets[loadedWorkbook.Worksheets.ActiveSheetIndex].ActiveCell = "C5";

            // Save the changes to a new file
            loadedWorkbook.Save("ActiveSheetDemo_Updated.xlsx");
        }
    }
}