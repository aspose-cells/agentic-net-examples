// Title: C# – Protect an Excel workbook with a password and recommend read‑only using Aspose.Cells
// Description: Shows how to enable write‑protection on a workbook, set a modification password and author, turn on the RecommendReadOnly flag, and save the file so Excel asks for the password only when a user attempts to edit.
// Keywords: Aspose.Cells C# protect workbook | Excel password modify only | write protection RecommendReadOnly | set workbook author Aspose | read‑only recommendation Excel | C# Excel file security Aspose.Cells
// Common Searches: Aspose.Cells set password to modify Excel file C# | How to enable RecommendReadOnly in Aspose.Cells | Add author to workbook protection Aspose.Cells | Create read‑only Excel workbook with password using .NET | C# code for write protection in Excel with Aspose
// Developer Intent: Add write‑protection with a modify password and suggest read‑only opening.
// Use Cases: Distribute a report that anyone can view but only authorized users can edit. | Provide a template that records the creator and encourages users to open it in read‑only mode. | Generate audit‑ready spreadsheets that prevent accidental changes without a password.
// AI Prompts: Generate C# code that applies write protection with a password and custom author to an existing workbook using Aspose.Cells. | Explain how to programmatically change or remove the RecommendReadOnly flag after a workbook has been saved. | Show an example of toggling the write‑protection password based on user input in a .NET console application.

using System;
using Aspose.Cells;

// Shows how to enable write‑protection on a workbook, set a modification password and author, turn on the RecommendReadOnly flag, and save the file so Excel asks for the password only when a user attempts to edit.
class ProtectWorkbookReadOnly
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the write‑protection settings
        WriteProtection writeProtection = workbook.Settings.WriteProtection;

        // Set the password required to modify the file
        writeProtection.Password = "modify123";

        // Optionally set the author of the protection
        writeProtection.Author = "Admin";

        // Recommend opening the file as read‑only
        writeProtection.RecommendReadOnly = true;

        // Save the workbook; users will be prompted for the password only if they try to edit
        workbook.Save("ReadOnlyProtected.xlsx", SaveFormat.Xlsx);
    }
}
