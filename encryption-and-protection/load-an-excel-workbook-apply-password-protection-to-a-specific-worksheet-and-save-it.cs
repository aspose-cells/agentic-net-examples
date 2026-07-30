// Title: C# – Password‑Protect a Single Worksheet with Aspose.Cells
// Description: Shows how to load an Excel workbook, apply full protection to a chosen worksheet using Aspose.Cells' Protect method, and save the file so the sheet is editable only with the specified password.
// Keywords: Aspose.Cells worksheet protection | C# protect Excel sheet password | Aspose.Cells Protect method | Excel sheet security .NET | password protect specific worksheet | load workbook Aspose.Cells | save protected workbook | Excel encryption Aspose
// Common Searches: Aspose.Cells protect worksheet C# example | How to set a password on a single Excel sheet using Aspose.Cells | C# code to lock a worksheet with a password Aspose | Save workbook after applying worksheet protection Aspose.Cells | Protect first sheet only Aspose.Cells .NET
// Developer Intent: Add password protection to a selected worksheet and persist the changes.
// Use Cases: Distribute a workbook while keeping confidential data on one sheet read‑only. | Create a template where only designated sheets can be edited by end users. | Enforce sheet‑level security in automated report generation pipelines.
// AI Prompts: Generate C# code that protects the second worksheet with password "Secret123" and saves as "protected.xlsx" using Aspose.Cells. | Explain how to programmatically unprotect a worksheet when the password is known, using Aspose.Cells for .NET. | Show how to apply different passwords to multiple worksheets in the same workbook with Aspose.Cells. | Provide a snippet to change the password of an already protected worksheet in an Excel file using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to load an Excel workbook, apply full protection to a chosen worksheet using Aspose.Cells' Protect method, and save the file so the sheet is editable only with the specified password.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        LoadOptions loadOptions = new LoadOptions(); // no password needed for loading in this example
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Select the worksheet you want to protect (e.g., the first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Protect the worksheet with a password.
        // ProtectionType.All protects all aspects of the sheet.
        // The third parameter (oldPassword) is null because the sheet is not yet protected.
        worksheet.Protect(ProtectionType.All, "MySheetPassword", null);

        // Save the workbook with the protected worksheet
        workbook.Save("output_protected.xlsx");
    }
}
