using System;
using Aspose.Cells;

class ProtectStructureRenameDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "password123");

        // Save the protected workbook
        string filePath = "protected_workbook.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the protected workbook
        Workbook loadedWorkbook = new Workbook(filePath);

        // Attempt to rename the first worksheet
        Worksheet sheet = loadedWorkbook.Worksheets[0];
        try
        {
            sheet.Name = "RenamedSheet";
            Console.WriteLine("Worksheet renamed unexpectedly (protection may not be enforced).");
        }
        catch (Exception ex)
        {
            // Expected: rename operation should be blocked due to structure protection
            Console.WriteLine("Rename operation blocked as expected: " + ex.Message);
        }

        // Verify that the workbook structure is still protected
        Console.WriteLine("Workbook structure protected: " + loadedWorkbook.Settings.IsProtected);
    }
}