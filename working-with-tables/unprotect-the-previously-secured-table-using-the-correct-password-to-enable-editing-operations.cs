// Title: Unprotect a password‑protected Excel worksheet with Aspose.Cells for .NET (C#) and edit its contents
// Common Searches: Aspose.Cells C# unprotect worksheet with password example | remove worksheet protection programmatically using Aspose.Cells .NET | how to edit a protected Excel sheet after unprotecting with Aspose.Cells | C# code to unprotect Excel table and modify a cell using Aspose.Cells | Aspose.Cells unprotect method usage for password‑protected workbook | unprotect Excel worksheet and save changes with Aspose.Cells for .NET
// Tags: Aspose.Cells unprotect worksheet password | C# remove Excel sheet protection Aspose.Cells | edit cell after worksheet unprotect Aspose.Cells | Aspose.Cells password protected workbook handling | unprotect Excel table using Aspose.Cells .NET

using Aspose.Cells;
using System;

// Loads an existing workbook, calls Worksheet.Unprotect with the correct password, modifies a cell, and saves the workbook, demonstrating how to remove protection from a worksheet/table and edit its data using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook that contains the protected table/worksheet
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that holds the table (adjust index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0]; // or workbook.Worksheets["Sheet1"]

        // Unprotect the worksheet using the correct password
        worksheet.Unprotect("YourPassword");

        // The worksheet is now editable; example of an edit operation
        worksheet.Cells["A1"].PutValue("Edited after unprotect");

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}
