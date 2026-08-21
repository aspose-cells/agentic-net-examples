// Title: Conditional Workbook Structure Protection (>10 Sheets) with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add multiple worksheets, evaluate the sheet count, and apply a password to the workbook's structure only when more than ten sheets are present, then save the file as XLSX.
// Keywords: Aspose.Cells conditional protection | C# protect workbook structure | password based on sheet count | Excel workbook security threshold | Aspose.Cells example protect if >10 sheets
// Common Searches: Aspose.Cells protect workbook only when sheet count exceeds 10 | C# conditional Excel password protection Aspose | apply structure password if workbook has many worksheets | how to lock Excel workbook based on number of sheets in .NET
// Developer Intent: Apply a password to the workbook's structure only when the workbook contains more than ten worksheets.
// Use Cases: Enforce policy that large workbooks (over ten tabs) must be locked to prevent accidental edits. | Automatically secure generated financial or reporting workbooks that exceed a sheet‑count threshold before distribution. | Implement dynamic protection in a macro‑driven project that creates variable numbers of worksheets.
// AI Prompts: Generate code that reads the protection password from a configuration file and applies it only when the worksheet count is greater than a configurable limit. | Show how to log a message indicating whether the workbook was protected based on its sheet count. | Explain how to protect both the workbook structure and windows conditionally, using a threshold defined in app settings.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add multiple worksheets, evaluate the sheet count, and apply a password to the workbook's structure only when more than ten sheets are present, then save the file as XLSX.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add worksheets so the workbook has more than ten sheets (for demonstration)
        for (int i = 0; i < 12; i++)
        {
            workbook.Worksheets.Add();
        }

        // Apply password protection only when the workbook contains more than ten worksheets
        if (workbook.Worksheets.Count > 10)
        {
            // Protect the workbook's structure with a password
            workbook.Protect(ProtectionType.Structure, "StrongPassword!123");
        }

        // Save the workbook to a file
        workbook.Save("ProtectedIfMoreThanTen.xlsx", SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}
