// Title: C# – Protect an Excel worksheet with Aspose.Cells while allowing only new hyperlinks
// Description: Demonstrates how to protect a worksheet using Aspose.Cells for .NET, enable insertion of new hyperlinks, block editing of existing cell content and hyperlink URLs, set a password, and save the workbook.
// Keywords: Aspose.Cells worksheet protection C# | allow hyperlink insertion Aspose.Cells | prevent editing existing hyperlinks | Excel sheet password protection .NET | Protection.AllowInsertingHyperlink example | Protection.AllowEditingContent usage
// Common Searches: Aspose.Cells protect sheet but allow adding hyperlinks | C# protect Excel worksheet and enable hyperlink insertion | block editing of existing hyperlink URLs Aspose.Cells | set password for worksheet protection Aspose.Cells .NET | how to allow only new hyperlinks on a protected Excel sheet
// Developer Intent: The developer needs a protected worksheet where users can add new hyperlinks but cannot modify any existing cell data or hyperlink addresses.
// Use Cases: Template workbook that lets users add reference links while keeping preset data immutable. | Financial report distributed to stakeholders, allowing them to insert source URLs without altering calculations. | Shared spreadsheet for documentation where only new hyperlinks are permitted, preserving original links.
// AI Prompts: Provide C# code using Aspose.Cells to protect a worksheet, enable only hyperlink insertion, and set a password. | Show an example that configures Protection.AllowInsertingHyperlink = true and Protection.AllowEditingContent = false. | Explain how to allow new hyperlinks on a protected Excel sheet while preventing changes to existing hyperlink URLs with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to protect a worksheet using Aspose.Cells for .NET, enable insertion of new hyperlinks, block editing of existing cell content and hyperlink URLs, set a password, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Configure protection: allow inserting new hyperlinks but prevent editing existing content (including hyperlink addresses)
        Protection protection = sheet.Protection;
        protection.AllowInsertingHyperlink = true;   // permit new hyperlink insertion
        protection.AllowEditingContent = false;     // block editing of existing cell contents / hyperlink addresses
        protection.Password = "pwd123";

        // Apply protection to the worksheet (all protection types)
        sheet.Protect(ProtectionType.All);

        // Demonstrate that inserting a hyperlink is allowed after protection
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

        // Save the workbook
        workbook.Save("ProtectedHyperlink.xlsx");
    }
}
