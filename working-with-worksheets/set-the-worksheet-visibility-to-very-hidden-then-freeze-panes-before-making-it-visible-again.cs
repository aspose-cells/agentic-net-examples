using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Add a temporary visible worksheet so the workbook always has at least one visible sheet
            Worksheet tempSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            tempSheet.Name = "Temp";

            // Access the original worksheet (the one we will hide temporarily)
            Worksheet sheet = workbook.Worksheets[0];

            // Set the worksheet to VeryHidden (allowed because another visible sheet exists)
            sheet.VisibilityType = VisibilityType.VeryHidden;

            // Freeze panes at cell C3 with 2 frozen rows and 2 frozen columns
            sheet.FreezePanes("C3", 2, 2);

            // Make the worksheet visible again
            sheet.VisibilityType = VisibilityType.Visible;

            // Remove the temporary sheet before saving (optional)
            workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);

            // Save the workbook
            workbook.Save("VeryHiddenFreezeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}