using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet named "Data"
        Worksheet dataWorksheet = workbook.Worksheets.Add("Data");
        int dataSheetIndex = dataWorksheet.Index;

        // Set the "Data" worksheet as the active sheet
        workbook.Worksheets.ActiveSheetIndex = dataSheetIndex;

        // Get the active worksheet
        Worksheet activeWorksheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Define the active cell within the active worksheet (e.g., C5)
        activeWorksheet.ActiveCell = "C5";

        // Optionally put a value in the active cell to verify
        activeWorksheet.Cells["C5"].PutValue("Active Cell");

        // Save the workbook
        workbook.Save("ActiveSheetAndCell.xlsx");

        // Load the saved workbook to verify the active sheet and cell
        Workbook loadedWorkbook = new Workbook("ActiveSheetAndCell.xlsx");
        Worksheet loadedActiveWorksheet = loadedWorkbook.Worksheets[loadedWorkbook.Worksheets.ActiveSheetIndex];
        Console.WriteLine("Active Sheet: " + loadedActiveWorksheet.Name);
        Console.WriteLine("Active Cell: " + loadedActiveWorksheet.ActiveCell);
    }
}