// Title: Apply Write Protection to an Aspose.Cells Workbook and Save in the Default XLSX Format (C#)
// Description: Creates a new Workbook, sets a password and a read‑only recommendation, saves it using the default XLSX format, then reloads the file to confirm the IsWriteProtected flag. Demonstrates how to prevent editing after save with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | write protection | password protected Excel | default XLSX save | RecommendReadOnly | IsWriteProtected | Excel file security | programmatic workbook protection
// Common Searches: Aspose.Cells set password for Excel workbook C# | Save a write‑protected XLSX with Aspose.Cells | Verify workbook protection after saving Aspose.Cells .NET | Enable read‑only recommendation in Aspose.Cells workbook | How to prevent editing of saved Excel file using Aspose.Cells
// Developer Intent: Add password‑based write protection to a workbook and store it in the default XLSX format.
// Use Cases: Distribute template files that require a password before any edits can be made. | Generate read‑only financial reports that enforce a password when users try to modify them. | Automate compliance checks by confirming that saved Excel files retain write protection.
// AI Prompts: Provide C# code that sets a password and RecommendReadOnly on an Aspose.Cells workbook, saves it as the default XLSX, and checks IsWriteProtected. | Show how to create a write‑protected Excel file with Aspose.Cells and verify the protection after loading the file. | Explain the steps to disable editing of a workbook saved in the default format using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a new Workbook, sets a password and a read‑only recommendation, saves it using the default XLSX format, then reloads the file to confirm the IsWriteProtected flag. Demonstrates how to prevent editing after save with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable write protection to prevent changes when saving
        workbook.Settings.WriteProtection.Password = "securePwd";
        workbook.Settings.WriteProtection.RecommendReadOnly = true; // recommend read‑only mode

        // Save the workbook using the default file format (XLSX)
        string outputPath = "ProtectedWorkbook.xlsx";
        workbook.Save(outputPath); // default SaveFormat is Xlsx

        // Load the saved workbook to verify the protection settings
        Workbook loadedWorkbook = new Workbook(outputPath);
        bool isWriteProtected = loadedWorkbook.Settings.WriteProtection.IsWriteProtected;
        Console.WriteLine("Workbook is write protected: " + isWriteProtected);
    }
}
