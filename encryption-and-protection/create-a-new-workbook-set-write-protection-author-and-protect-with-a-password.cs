// Title: Aspose.Cells C# – Create Workbook with Write‑Protection Author & Password
// Description: Demonstrates how to instantiate a new Workbook, assign a write‑protection author, set a password, optionally recommend read‑only mode, and save the file as a protected Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells write protection | C# protect Excel workbook | set write‑protection author | Excel password protection .NET | RecommendReadOnly Aspose.Cells | save protected workbook
// Common Searches: Aspose.Cells set workbook write protection author C# | protect Excel file with password using Aspose.Cells | recommend read‑only mode Aspose.Cells workbook | how to add write protection to a new Excel workbook .NET | C# example write protection author Aspose.Cells
// Developer Intent: Create a fresh Excel workbook and apply write‑protection with a specific author and password, optionally marking it as read‑only.
// Use Cases: Distribute templates that users can view but not edit, with author attribution for audit trails. | Secure financial or regulatory reports against accidental changes while suggesting read‑only opening. | Provide configuration spreadsheets for internal tools that must remain unchanged unless authorized.
// AI Prompts: Generate C# code with Aspose.Cells to create a workbook, set WriteProtection.Author to "John Doe", set WriteProtection.Password to "password123", enable RecommendReadOnly, and save as "WriteProtectedWorkbook.xlsx". | Explain the effect of Aspose.Cells WriteProtection settings on Excel behavior and show how to modify or remove the protection programmatically.

using System;
using Aspose.Cells;

// Demonstrates how to instantiate a new Workbook, assign a write‑protection author, set a password, optionally recommend read‑only mode, and save the file as a protected Excel workbook using Aspose.Cells for .NET.
class WriteProtectionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the author for write protection
        workbook.Settings.WriteProtection.Author = "John Doe";

        // Set the password that protects the workbook from modification
        workbook.Settings.WriteProtection.Password = "password123";

        // (Optional) Recommend opening the file as read‑only
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook with the write‑protection settings applied
        workbook.Save("WriteProtectedWorkbook.xlsx");
    }
}
