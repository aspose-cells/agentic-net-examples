using System;
using Aspose.Cells;

class ShiftWorksheetToEnd
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains a default sheet)
            Workbook workbook = new Workbook();

            // Remove the default sheet to avoid duplicate names
            if (workbook.Worksheets.Count > 0)
                workbook.Worksheets.RemoveAt(0);

            // Add sample worksheets with unique names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Choose the worksheet to move (e.g., "Sheet2")
            Worksheet sheetToMove = workbook.Worksheets["Sheet2"];

            // Destination index is the last position (zero‑based)
            int lastIndex = workbook.Worksheets.Count - 1;

            // Move the selected worksheet to the end of the workbook
            sheetToMove.MoveTo(lastIndex);

            // Save the workbook
            workbook.Save("ShiftedWorksheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}