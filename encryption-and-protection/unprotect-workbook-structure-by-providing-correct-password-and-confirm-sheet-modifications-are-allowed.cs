// Title: Unprotect Excel workbook structure with password using Aspose.Cells for .NET
// Description: Demonstrates how to protect a workbook's structure, save it, reload it, verify the protection flag, remove the structure protection with the correct password, add a new worksheet, and save the unprotected file using Aspose.Cells in C#.
// Keywords: Aspose.Cells unprotect workbook | C# remove workbook structure protection | IsWorkbookProtectedWithPassword | Excel workbook password protection .NET | add worksheet after unprotect
// Common Searches: Aspose.Cells unprotect workbook structure C# | check workbook protection status Aspose.Cells | add sheet after removing Excel protection .NET | how to use Unprotect method with password
// Developer Intent: Programmatically remove structure protection from an Excel workbook by providing the correct password and verify that further sheet modifications are allowed.
// Use Cases: Load a password‑protected workbook, call Unprotect with the password, and save the editable file. | Read the IsWorkbookProtectedWithPassword property before and after unprotecting to confirm the protection state. | Automate a process that initially locks a workbook for distribution and later unlocks it for data manipulation.
// AI Prompts: Write C# code with Aspose.Cells that unprotects a workbook's structure using a given password, adds a new worksheet, and saves the result. | Explain the purpose of IsWorkbookProtectedWithPassword and how to use it to validate protection status around an Unprotect call. | Show error‑handling patterns for an incorrect password when calling Workbook.Unprotect in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to protect a workbook's structure, save it, reload it, verify the protection flag, remove the structure protection with the correct password, add a new worksheet, and save the unprotected file using Aspose.Cells in C#.
    class UnprotectWorkbookStructureDemo
    {
        static void Main()
        {
            // Create a new workbook and protect its structure with a password
            Workbook wb = new Workbook();
            wb.Protect(ProtectionType.Structure, "mySecretPwd");
            // Save the protected workbook
            wb.Save("protected_workbook.xlsx");

            // Load the protected workbook
            Workbook protectedWb = new Workbook("protected_workbook.xlsx");

            // Verify that the workbook is indeed protected with a password
            Console.WriteLine("Is workbook protected with password? " + protectedWb.IsWorkbookProtectedWithPassword);

            // Unprotect the workbook structure using the correct password
            protectedWb.Unprotect("mySecretPwd");

            // Confirm that the workbook is no longer protected
            Console.WriteLine("Is workbook still protected? " + protectedWb.IsWorkbookProtectedWithPassword);

            // Perform a modification to ensure changes are now allowed (e.g., add a new worksheet)
            Worksheet newSheet = protectedWb.Worksheets.Add("NewSheet");
            newSheet.Cells["A1"].PutValue("Modification after unprotect");

            // Save the unprotected workbook
            protectedWb.Save("unprotected_workbook.xlsx");

            Console.WriteLine("Workbook unprotected and modified successfully.");
        }
    }
}
