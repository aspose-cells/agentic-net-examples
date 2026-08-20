// Title: List password‑protected worksheets in an Excel file using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook with Aspose.Cells, scans each worksheet, checks the Worksheet.Protection.IsProtectedWithPassword flag, gathers the names of sheets secured with a password, and prints the list to the console.
// Keywords: Aspose.Cells list protected worksheets | C# detect worksheet password protection | enumerate password‑protected sheets Aspose | Excel workbook protection status .NET | retrieve protected sheet names Aspose.Cells | global Excel security audit | Aspose.Cells .NET worldwide
// Common Searches: how to list password protected worksheets using Aspose.Cells | C# code to check Excel sheet protection Aspose | get names of protected worksheets in a workbook .NET | identify password‑protected sheets with Aspose.Cells
// Developer Intent: Identify and output all worksheets in a workbook that are secured with a password.
// Use Cases: Generate a quick report of protected sheets before sharing the workbook. | Validate that only intended worksheets are password‑protected for compliance. | Log protected worksheet names for security audits or change‑management processes.
// AI Prompts: Provide C# code that writes the list of password‑protected worksheets to a CSV file using Aspose.Cells. | Show how to programmatically remove password protection from the identified worksheets. | Explain how to check both worksheet and workbook protection status in a single Aspose.Cells routine.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, scans each worksheet, checks the Worksheet.Protection.IsProtectedWithPassword flag, gathers the names of sheets secured with a password, and prints the list to the console.
class ListProtectedWorksheets
{
    static void Main()
    {
        // Path to the workbook to be examined
        string filePath = "input.xlsx";

        // Load the workbook (uses Aspose.Cells default load options)
        Workbook workbook = new Workbook(filePath);

        // Collect names of worksheets that are password protected
        List<string> protectedSheets = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Worksheet.Protection.IsProtectedWithPassword indicates password protection
            if (sheet.Protection.IsProtectedWithPassword)
            {
                protectedSheets.Add(sheet.Name);
            }
        }

        // Display the result
        Console.WriteLine("Password protected worksheets:");
        foreach (string name in protectedSheets)
        {
            Console.WriteLine("- " + name);
        }
    }
}
