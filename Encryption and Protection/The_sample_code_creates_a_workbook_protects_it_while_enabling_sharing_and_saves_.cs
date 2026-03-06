using System;
using Aspose.Cells;

class ProtectSharedWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Enable sharing for the workbook
        wb.Settings.Shared = true;

        // Protect the shared workbook with a password
        wb.ProtectSharedWorkbook("myPassword");

        // Save the protected shared workbook as XLSX
        string outputPath = "ProtectedSharedWorkbook.xlsx";
        wb.Save(outputPath, SaveFormat.Xlsx);

        // Load the saved workbook to verify protection and sharing status
        Workbook loadedWb = new Workbook(outputPath);

        // Output verification results
        Console.WriteLine("Workbook is shared: " + loadedWb.Settings.Shared);
        Console.WriteLine("Workbook is protected (structure/window): " + loadedWb.Settings.IsProtected);
        Console.WriteLine("Workbook is protected with password: " + loadedWb.IsWorkbookProtectedWithPassword);
    }
}