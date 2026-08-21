// Title: C# – Protect a Single Worksheet with a Password Using Aspose.Cells for .NET
// Description: Load an existing Excel workbook, assign a password to a chosen worksheet, apply full protection (ProtectionType.All), and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | worksheet protection | password protection | Excel encryption .NET | Protect worksheet Aspose | ProtectionType.All | load workbook Aspose | save protected workbook | Excel sheet security
// Common Searches: Aspose.Cells protect worksheet password C# | How to set password for a single sheet using Aspose.Cells | C# code to apply ProtectionType.All to Excel sheet | Save Excel workbook with protected sheet Aspose | Encrypt specific worksheet Aspose.Cells .NET
// Developer Intent: Load an Excel file, secure a specific worksheet with a password, and write the protected workbook back to disk.
// Use Cases: Distribute a workbook while keeping confidential data hidden on a protected sheet. | Lock formula cells in a reporting sheet while allowing other sheets to remain editable. | Enforce read‑only access to a financial model worksheet for external reviewers.
// AI Prompts: Generate C# code that protects the second worksheet with a password and permits only formatting changes using Aspose.Cells. | Show how to detect if a worksheet is already protected before applying a new password in Aspose.Cells for .NET. | Provide an example that protects multiple worksheets, each with a different password, in a single workbook using Aspose.Cells.

using System;
using Aspose.Cells;

// Load an existing Excel workbook, assign a password to a chosen worksheet, apply full protection (ProtectionType.All), and save the file with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook (no password needed for loading in this example)
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the worksheet you want to protect (e.g., the first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the password for the worksheet protection
        worksheet.Protection.Password = "MySheetPassword";

        // Apply protection to the worksheet with the specified password
        // ProtectionType.All protects all aspects of the worksheet
        worksheet.Protect(ProtectionType.All, "MySheetPassword", null);

        // Save the workbook with the protected worksheet
        workbook.Save("output_protected.xlsx");
    }
}
