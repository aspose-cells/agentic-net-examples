using System;
using Aspose.Cells;

class RelocateWorksheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains a default sheet named "Sheet1")
            Workbook workbook = new Workbook();

            // Remove the default sheet to avoid duplicate names
            workbook.Worksheets.Clear();

            // Add worksheets in the desired order: Sheet1, TargetSheet, Sheet3
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("TargetSheet");
            workbook.Worksheets.Add("Sheet3");

            // Retrieve the worksheet that needs to be moved to the first position
            Worksheet targetSheet = workbook.Worksheets["TargetSheet"];

            // Move the worksheet to index 0 (the first tab)
            targetSheet.MoveTo(0);

            // Ensure the first visible tab points to the moved sheet
            workbook.Settings.FirstVisibleTab = 0;

            // Save the workbook with the updated sheet order
            workbook.Save("RelocatedWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}