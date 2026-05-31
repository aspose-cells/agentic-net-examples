using System;
using Aspose.Cells;

class WorkbookProtectionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "myPassword");

        // Verify that the workbook is protected with a password
        Console.WriteLine("Is workbook protected with password: " + workbook.IsWorkbookProtectedWithPassword);

        // Attempt to add a new worksheet (should be blocked by protection)
        try
        {
            workbook.Worksheets.Add("NewSheet");
            Console.WriteLine("Worksheet added successfully (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to add worksheet as expected: " + ex.Message);
        }

        // Save the protected workbook
        string filePath = "ProtectedWorkbook.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);
        workbook.Dispose();

        // Load the saved workbook to verify protection persists
        Workbook loadedWorkbook = new Workbook(filePath);
        Console.WriteLine("Loaded workbook protected with password: " + loadedWorkbook.IsWorkbookProtectedWithPassword);

        // Attempt to add another worksheet after loading (should also fail)
        try
        {
            loadedWorkbook.Worksheets.Add("AnotherSheet");
            Console.WriteLine("Worksheet added after load (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to add worksheet after load as expected: " + ex.Message);
        }

        loadedWorkbook.Dispose();
    }
}