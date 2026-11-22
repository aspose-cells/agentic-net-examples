// Title: Add a worksheet to a password‑protected Excel workbook and handle a failed unprotect attempt with an incorrect password using Aspose.Cells for .NET
// Common Searches: how to add a new sheet to a password protected Excel file with Aspose.Cells C# | catch exception when unprotecting a workbook with wrong password using Aspose.Cells | Aspose.Cells example adding worksheet to protected.xlsx without removing protection | save workbook after failed unprotect attempt Aspose.Cells .NET | C# code to load protected workbook, add sheet, and test invalid unprotect password
// Tags: add worksheet to protected workbook Aspose.Cells | unprotect workbook with invalid password C# | exception handling for workbook unprotect Aspose.Cells | save workbook after protection error Aspose.Cells | Aspose.Cells workbook protection workflow

using Aspose.Cells;
using System;

// The sample loads a password‑protected Excel file (protected.xlsx), adds a new worksheet named "NewSheet", then attempts to unprotect the workbook using an incorrect password inside a try‑catch block, logs the exception message, and finally saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing protected workbook
        Workbook workbook = new Workbook("protected.xlsx");

        // Add a new worksheet to the workbook
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet newSheet = workbook.Worksheets[newSheetIndex];
        newSheet.Name = "NewSheet";

        // Attempt to unprotect the workbook with an incorrect password
        try
        {
            workbook.Unprotect("wrongPassword");
        }
        catch (Exception ex)
        {
            // Expected failure because the password is incorrect
            Console.WriteLine("Unprotect failed: " + ex.Message);
        }

        // Save the workbook after the operation
        workbook.Save("output.xlsx");
    }
}
