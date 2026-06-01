using System;
using Aspose.Cells;

class ProtectWorkbookStructureDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add extra worksheets for demonstration
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "pwd123");

        // Verify that the workbook structure is protected
        Console.WriteLine("Workbook structure protected: " + workbook.Settings.IsProtected);
        Console.WriteLine("Workbook protected with password: " + workbook.IsWorkbookProtectedWithPassword);

        // Attempt to delete a worksheet (should be blocked)
        try
        {
            // Try to remove the second worksheet (index 1)
            workbook.Worksheets.RemoveAt(1);
            Console.WriteLine("Worksheet removed successfully (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to remove worksheet as expected: " + ex.Message);
        }

        // Save the workbook
        workbook.Save("ProtectedStructureWorkbook.xlsx", SaveFormat.Xlsx);
    }
}