// Title: C# – Unlock a password‑protected VBA project in an XLSM file with Aspose.Cells
// Description: Loads a macro‑enabled workbook, accesses its VbaProject, checks the protection flag, validates a supplied password using ValidatePassword, removes protection with Protect(false, null) when the password is correct, and saves the workbook as an unlocked XLSM file.
// Keywords: Aspose.Cells VBA unlock | C# remove VBA password | Validate VBA project password | Unprotect VBA project programmatically | Save unlocked XLSM workbook
// Common Searches: how to unlock VBA project programmatically c# | aspocells validate vba password | remove password from macro enabled workbook aspocells | c# unprotect vba project in xlsm | aspocells vba project protect false null
// Developer Intent: Provide a C# example that unlocks a VBA project protected for viewing by validating a password and then removing the protection, finally saving the workbook without VBA protection.
// Use Cases: Detect whether a loaded workbook contains a VBA project and report its initial protection status. | Validate a user‑supplied password against the VBA project and, if correct, clear the protection. | Persist the unprotected workbook by saving it to a new XLSM file.
// AI Prompts: Write C# code using Aspose.Cells to check a VBA project's protection state, validate a password, and unprotect it. | Suggest robust error‑handling patterns for unlocking a VBA project in a console application with Aspose.Cells. | Explain the difference between VbaProject.IsProtected and VbaProject.Protect parameters when removing VBA protection.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaUnlockDemo
{
    // Loads a macro‑enabled workbook, accesses its VbaProject, checks the protection flag, validates a supplied password using ValidatePassword, removes protection with Protect(false, null) when the password is correct, and saves the workbook as an unlocked XLSM file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro‑enabled workbook that has a VBA project locked for viewing
            string workbookPath = "LockedVbaProject.xlsm";

            // Password that is supposed to unlock the VBA project
            string vbaPassword = "yourPassword";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input file not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Report initial protection state
                Console.WriteLine($"Initial VBA Project IsProtected: {vbaProject.IsProtected}");

                // Validate the provided password
                bool isPasswordValid = vbaProject.ValidatePassword(vbaPassword);
                Console.WriteLine($"Password validation result: {isPasswordValid}");

                if (isPasswordValid)
                {
                    // Unprotect the VBA project (remove protection)
                    // Passing false for isProtected and null for password removes protection
                    vbaProject.Protect(false, null);
                    Console.WriteLine("VBA project has been successfully unprotected.");
                }
                else
                {
                    Console.WriteLine("Failed to unprotect VBA project: invalid password.");
                }

                // Report final protection state
                Console.WriteLine($"Final VBA Project IsProtected: {vbaProject.IsProtected}");

                // Optionally save the workbook to a new file to persist changes
                string outputPath = "UnlockedVbaProject.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
