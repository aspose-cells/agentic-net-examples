// Title: Protect Excel Workbook Structure with a Password Using Aspose.Cells for .NET
// Description: Creates a Workbook, applies structure‑only protection with a password via Workbook.Protect, checks the IsWorkbookProtectedWithPassword flag before and after saving, and confirms the protection persists when the file is reloaded.
// Keywords: Aspose.Cells | .NET | C# | Workbook.Protect | structure protection | password protection | IsWorkbookProtectedWithPassword | Excel security | save and reload verification | prevent sheet addition
// Common Searches: Aspose.Cells protect workbook structure password | check if Excel workbook is password protected after save .NET | verify workbook protection after loading file | C# protect Excel sheet layout with Aspose.Cells | prevent adding or deleting worksheets using Aspose.Cells
// Developer Intent: Apply a password to lock the workbook’s structure and ensure the protection remains after the file is saved and reopened.
// Use Cases: Lock template workbooks so users can only edit data, not modify sheet order. | Automate compliance checks that exported reports retain structure protection before distribution. | Restrict end‑users from adding, removing, or renaming worksheets in generated Excel files.
// AI Prompts: Write C# code with Aspose.Cells to protect only the workbook structure using a password and validate the protection after saving. | Show how to handle exceptions when trying to modify a structure‑protected workbook with Aspose.Cells. | Explain how to change or remove workbook structure protection and update the password in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a Workbook, applies structure‑only protection with a password via Workbook.Protect, checks the IsWorkbookProtectedWithPassword flag before and after saving, and confirms the protection persists when the file is reloaded.
class ProtectWorkbookStructure
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "mySecretPassword");

        // Verify that the workbook is now protected with a password
        Console.WriteLine("Is workbook protected with password (before save): " + workbook.IsWorkbookProtectedWithPassword);

        // Save the protected workbook
        string filePath = "ProtectedWorkbook.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the saved workbook to verify protection persists
        Workbook loadedWorkbook = new Workbook(filePath);

        // Verify protection status of the loaded workbook
        Console.WriteLine("Is loaded workbook protected with password: " + loadedWorkbook.IsWorkbookProtectedWithPassword);

        // Clean up resources
        workbook.Dispose();
        loadedWorkbook.Dispose();
    }
}
