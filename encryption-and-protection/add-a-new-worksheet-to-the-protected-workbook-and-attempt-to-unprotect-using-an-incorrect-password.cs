// Title: Add a worksheet to a password‑protected Excel file and handle an invalid unprotect password with Aspose.Cells for .NET
// AI Prompts: Load a protected Excel workbook, insert a worksheet named "NewSheet", call Workbook.Unprotect with an incorrect password inside a try‑catch block, and save the result using Aspose.Cells in C#. | Show how to catch the exception thrown when Workbook.Unprotect receives a wrong password after adding a new sheet to a protected workbook with Aspose.Cells for .NET. | Demonstrate error handling for a failed workbook unprotect operation after adding a worksheet to a password‑protected Excel file using the Aspose.Cells C# API.
// Common Searches: asp.net aspose.cells add worksheet to encrypted Excel file | c# workbook.unprotect wrong password exception handling | how to catch unprotect error when workbook is password protected aspose.cells
// Tags: add worksheet to protected workbook Aspose.Cells | unprotect workbook with incorrect password C# | exception handling for Workbook.Unprotect Aspose.Cells | save modified protected Excel file Aspose.Cells | load password protected Excel Aspose.Cells .NET

using Aspose.Cells;
using System;

// // Loads a password‑protected Excel workbook, adds a new worksheet called "NewSheet", attempts to unprotect the workbook with an invalid password inside a try‑catch block, logs the failure, and saves the workbook as "ModifiedWorkbook.xlsx".
class Program
{
    static void Main()
    {
        // Load the existing protected workbook
        Workbook workbook = new Workbook("ProtectedWorkbook.xlsx"); // load rule

        // Add a new worksheet to the workbook
        workbook.Worksheets.Add("NewSheet"); // create worksheet

        // Attempt to unprotect the workbook using an incorrect password
        try
        {
            workbook.Unprotect("WrongPassword"); // attempt unprotect
            Console.WriteLine("Workbook unprotected (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to unprotect workbook: " + ex.Message);
        }

        // Save the modified workbook
        workbook.Save("ModifiedWorkbook.xlsx"); // save rule
    }
}
