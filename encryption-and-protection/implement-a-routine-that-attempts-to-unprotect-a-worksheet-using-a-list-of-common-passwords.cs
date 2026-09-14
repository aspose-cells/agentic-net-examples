// Title: Attempt to unprotect an Excel worksheet by trying common passwords with Aspose.Cells for .NET
// AI Prompts: Write a C# method that receives a Worksheet object and iterates over a predefined list of typical passwords, calling Worksheet.Unprotect for each until the sheet is no longer protected. | Modify the unprotect routine to return the password that succeeded or null if none of the common passwords work. | Add detailed logging to the password‑checking loop, capturing each attempted password and any CellsException thrown.
// Common Searches: how to programmatically unprotect an Excel worksheet with Aspose.Cells using a list of passwords in C# | c# Aspose.Cells try multiple passwords to remove worksheet protection | sample code for iterating common passwords with Worksheet.Unprotect in .NET | catch CellsException while attempting to unprotect a protected sheet using Aspose.Cells
// Tags: worksheet.unprotect with password list Aspose.Cells | c# iterate common passwords Excel protection | handle cellsexception Aspose.Cells unprotect | bulk password trial for Excel sheet protection .NET | unprotect worksheet programmatically Aspose.Cells

using System;
using Aspose.Cells;

// Loads 'ProtectedWorkbook.xlsx', iterates through a set of common passwords calling Worksheet.Unprotect until the sheet is no longer protected, then saves the workbook as 'UnprotectedWorkbook.xlsx'.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("ProtectedWorkbook.xlsx");

        // Get the first worksheet (or any specific worksheet you want to unprotect)
        Worksheet sheet = workbook.Worksheets[0];

        // Attempt to unprotect the worksheet using common passwords
        TryUnprotectWorksheet(sheet);

        // Save the workbook after unprotection attempt
        workbook.Save("UnprotectedWorkbook.xlsx");
    }

    /// <param name="sheet">The worksheet to unprotect.</param>
    static void TryUnprotectWorksheet(Worksheet sheet)
    {
        // List of common passwords to try
        string[] commonPasswords = new string[]
        {
            "password",
            "1234",
            "admin",
            "123456",
            "letmein",
            "password1",
            "abc123",
            "qwerty",
            "1111",
            "12345"
        };

        // If the worksheet is already unprotected, no need to try passwords
        if (!sheet.IsProtected)
            return;

        foreach (string pwd in commonPasswords)
        {
            try
            {
                // Attempt to unprotect with the current password
                sheet.Unprotect(pwd);

                // If no exception was thrown, check if unprotected succeeded
                if (!sheet.IsProtected)
                {
                    // Successfully unprotected; exit the loop
                    break;
                }
            }
            catch (CellsException)
            {
                // Wrong password – continue with the next one
                continue;
            }
        }
    }
}
