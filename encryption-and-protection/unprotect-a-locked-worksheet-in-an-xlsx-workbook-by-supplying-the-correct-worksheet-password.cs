// Title: How to unprotect a specific worksheet in a password‑protected XLSX file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens a password‑protected .xlsx workbook with Aspose.Cells, calls Worksheet.Unprotect with the known password, and saves the result as an unprotected file. | Show how to select a worksheet by name or index and remove its protection using Aspose.Cells' Unprotect method in a .NET application.
// Common Searches: aspnet unprotect worksheet with known password using Aspose.Cells | c# remove sheet protection from protected.xlsx using Aspose.Cells library | example code to unprotect a single sheet in an Excel workbook with Aspose.Cells for .NET | how to load a protected Excel file and save it without sheet password in C#
// Tags: Aspose.Cells worksheet unprotect C# | remove Excel sheet password programmatically | load protected XLSX save unprotected Aspose.Cells | C# unprotect specific worksheet Aspose.Cells

using System;
using Aspose.Cells;

// Loads a password‑protected XLSX workbook, removes protection from a chosen worksheet using the supplied password via Worksheet.Unprotect, and saves the workbook as an unprotected file.
class Program
{
    static void Main()
    {
        // Load the protected workbook from file
        Workbook workbook = new Workbook("protected.xlsx");

        // Get the worksheet you want to unprotect (by name or index)
        Worksheet sheet = workbook.Worksheets["Sheet1"]; // replace with actual sheet name or use Worksheets[0]

        // Unprotect the worksheet using the known password
        sheet.Unprotect("YourPasswordHere"); // replace with the correct password

        // Save the workbook with the worksheet now unprotected
        workbook.Save("unprotected.xlsx");
    }
}
