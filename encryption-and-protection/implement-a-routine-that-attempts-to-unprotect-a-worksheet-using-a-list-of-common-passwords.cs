// Title: Programmatically unprotect Excel worksheets by testing common passwords with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to try multiple passwords for protected Excel worksheets | how to automatically unprotect all sheets in an .xlsx using a password list with Aspose.Cells | sample C# routine to brute‑force Excel worksheet protection using common passwords | unprotect protected workbook worksheets in .NET without knowing the password | Aspose.Cells TryUnprotectWorksheet method example with password collection
// Tags: Aspose.Cells unprotect worksheet password list | C# batch unprotect Excel worksheets | handle CellsException during worksheet unprotect | iterate workbook worksheets Aspose.Cells | common passwords for Excel sheet protection

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsUnprotectDemo
{
    // Loads a workbook, defines a collection of typical passwords, and iterates through each worksheet. For each sheet it checks the protection flag and attempts to unprotect using the supplied passwords, catching CellsException for incorrect attempts. The routine reports success per sheet and saves the workbook as an unprotected copy.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("ProtectedWorkbook.xlsx");

            // List of common passwords to try
            List<string> commonPasswords = new List<string>
            {
                "password",
                "123456",
                "admin",
                "letmein",
                "welcome",
                "1234",
                "12345",
                "12345678",
                "qwerty",
                "abc123"
            };

            // Attempt to unprotect each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool unprotected = TryUnprotectWorksheet(sheet, commonPasswords);
                Console.WriteLine($"Worksheet \"{sheet.Name}\" unprotected: {unprotected}");
            }

            // Save the workbook after attempting unprotection
            workbook.Save("UnprotectedWorkbook.xlsx");
        }

        /// <param name="sheet">The worksheet to unprotect.</param>
        /// <param name="passwords">A collection of passwords to try.</param>
        /// <returns>True if unprotected; otherwise false.</returns>
        static bool TryUnprotectWorksheet(Worksheet sheet, IEnumerable<string> passwords)
        {
            // If the worksheet is not protected, nothing to do.
            if (!sheet.IsProtected)
                return true;

            foreach (string pwd in passwords)
            {
                try
                {
                    // Attempt to unprotect with the current password.
                    sheet.Unprotect(pwd);

                    // If no exception, the password was correct.
                    return true;
                }
                catch (CellsException)
                {
                    // Incorrect password – continue to the next one.
                }
            }

            // None of the passwords succeeded.
            return false;
        }
    }
}
