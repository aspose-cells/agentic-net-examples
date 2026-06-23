using System;
using Aspose.Cells;

class ProtectWorkbookStructureDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "mySecretPwd");

        // Verify protection status before saving
        Console.WriteLine("IsWorkbookProtectedWithPassword (before save): " + workbook.IsWorkbookProtectedWithPassword);
        Console.WriteLine("Settings.IsProtected (before save): " + workbook.Settings.IsProtected);

        // Save the protected workbook
        string filePath = "ProtectedWorkbook.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the saved workbook
        Workbook loadedWorkbook = new Workbook(filePath);

        // Verify protection status after loading
        Console.WriteLine("IsWorkbookProtectedWithPassword (after load): " + loadedWorkbook.IsWorkbookProtectedWithPassword);
        Console.WriteLine("Settings.IsProtected (after load): " + loadedWorkbook.Settings.IsProtected);

        // Dispose objects
        workbook.Dispose();
        loadedWorkbook.Dispose();
    }
}