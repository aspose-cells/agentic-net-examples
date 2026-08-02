// Title: C# – Protect Excel Workbook Structure with a Password Using Aspose.Cells and Verify It
// Description: This example shows how to create a workbook, lock its structure using a password via Workbook.Protect, examine the in‑memory protection flags (IsWorkbookProtectedWithPassword and Settings.IsProtected), save the file, reopen it, and confirm that the lock remains active.
// Keywords: Aspose.Cells | C# workbook protection | Excel structure lock | password protected workbook | IsWorkbookProtectedWithPassword | Workbook.Settings.IsProtected | prevent sheet deletion | save and reload protection | ProtectionType.Structure | Excel security .NET
// Common Searches: Aspose.Cells protect workbook structure C# | Check workbook protection after saving Aspose.Cells | How to lock sheet tabs with password in .NET | Verify Excel workbook structure lock after reload | Prevent adding worksheets using Aspose.Cells
// Developer Intent: Add a password‑based structure lock to an Excel file and ensure the lock persists after the file is saved and reopened.
// Use Cases: Distribute a template that must keep its original worksheets | Create a read‑only report where users cannot modify sheet layout | Automate compliance by enforcing workbook structure in generated files | Programmatically confirm protection before sending files to clients
// AI Prompts: Generate C# code that sets a password on the workbook structure with Aspose.Cells, saves the file, reloads it, and prints the protection status. | Show how to read IsWorkbookProtectedWithPassword and Settings.IsProtected after opening a protected Excel workbook. | Explain the steps to apply ProtectionType.Structure and verify persistence across save/load in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// This example shows how to create a workbook, lock its structure using a password via Workbook.Protect, examine the in‑memory protection flags (IsWorkbookProtectedWithPassword and Settings.IsProtected), save the file, reopen it, and confirm that the lock remains active.
class ProtectWorkbookStructureDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "mySecretPwd");

        // Verify protection in memory
        Console.WriteLine("Is workbook protected with password (in-memory): " + workbook.IsWorkbookProtectedWithPassword);
        Console.WriteLine("WorkbookSettings.IsProtected (in-memory): " + workbook.Settings.IsProtected);

        // Save the protected workbook
        string filePath = "ProtectedStructureWorkbook.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);
        workbook.Dispose();

        // Load the saved workbook
        Workbook loadedWorkbook = new Workbook(filePath);

        // Verify protection after loading
        Console.WriteLine("Is loaded workbook protected with password: " + loadedWorkbook.IsWorkbookProtectedWithPassword);
        Console.WriteLine("Loaded WorkbookSettings.IsProtected: " + loadedWorkbook.Settings.IsProtected);

        loadedWorkbook.Dispose();
    }
}
