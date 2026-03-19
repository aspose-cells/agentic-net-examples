using System;
using Aspose.Cells;

namespace ActivateWorksheetDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with the default worksheet
            Workbook workbook = new Workbook();

            // Add a second worksheet named "DataSheet"
            Worksheet dataSheet = workbook.Worksheets.Add("DataSheet");

            // Set the second worksheet as the active sheet when the workbook is opened
            workbook.Worksheets.ActiveSheetIndex = dataSheet.Index;
            dataSheet.IsSelected = true; // Ensure the sheet is selected

            // Set a specific cell (e.g., "C3") as the active cell in the active worksheet
            dataSheet.ActiveCell = "C3";

            // Optionally, you can also select a range that includes the active cell
            // dataSheet.SelectRange(2, 2, 1, 1, true); // Row 2, Column 2 corresponds to "C3"

            // Save the workbook to a file
            string outputPath = "ActivatedWorksheet.xlsx";
            workbook.Save(outputPath);

            // Load the saved workbook to verify the active sheet and cell
            Workbook loadedWorkbook = new Workbook(outputPath);
            Worksheet activeWorksheet = loadedWorkbook.Worksheets[loadedWorkbook.Worksheets.ActiveSheetIndex];
            Console.WriteLine("Active Worksheet: " + activeWorksheet.Name);
            Console.WriteLine("Active Cell: " + activeWorksheet.ActiveCell);
        }
    }
}