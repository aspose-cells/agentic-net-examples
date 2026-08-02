// Title: C# – Handle Empty Password Exception for Aspose.Cells VbaProject.Protect
// Description: Shows how to create a Workbook, protect its VBA project, catch the exception thrown when an empty password is used, then protect with a valid password and save the macro‑enabled file, applying error handling throughout.
// Keywords: Aspose.Cells | VbaProject.Protect | empty password exception | C# error handling | macro‑enabled workbook | VBA project protection | .NET | try‑catch | Aspose.Cells VBA
// Common Searches: Aspose.Cells VbaProject.Protect empty password error | C# catch exception when protecting VBA project with Aspose.Cells | how to handle empty password in VbaProject.Protect | protect macro workbook Aspose.Cells C# example | error handling for VBA project protection Aspose
// Developer Intent: Detect and manage the exception raised by VbaProject.Protect when the password argument is empty.
// Use Cases: Validate password input before calling Protect to avoid runtime failures. | Log detailed exception data for diagnostics when protection cannot be applied. | Continue processing or abort saving based on the success of the protection step.
// AI Prompts: Create a reusable C# wrapper for workbook.VbaProject.Protect that returns a boolean and logs errors for empty passwords. | Write NUnit tests that verify exception handling for VbaProject.Protect with both empty and non‑empty passwords. | Generate C# code that captures the stack trace of a Protect failure and falls back to saving the workbook without protection.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    // Shows how to create a Workbook, protect its VBA project, catch the exception thrown when an empty password is used, then protect with a valid password and save the macro‑enabled file, applying error handling throughout.
    public class VbaProjectProtectWithErrorHandling
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (this automatically creates a VBA project container)
            Workbook workbook = new Workbook();

            // Attempt to protect the VBA project with an empty password.
            // According to the API, providing an empty password when locking for viewing
            // may raise an exception. We capture any exception here.
            try
            {
                // isLockedForViewing = false, password = empty string
                workbook.VbaProject.Protect(false, string.Empty);
                Console.WriteLine("VBA project protected successfully with empty password (no view lock).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception caught while protecting VBA project with empty password: {ex.Message}");
            }

            // For comparison, protect the VBA project with a valid password and view lock.
            try
            {
                workbook.VbaProject.Protect(true, "StrongPassword123");
                Console.WriteLine("VBA project protected and locked for viewing with a valid password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception caught while protecting VBA project with a valid password: {ex.Message}");
            }

            // Save the workbook as a macro‑enabled file to persist the VBA project.
            try
            {
                workbook.Save("VbaProjectProtected.xlsm", SaveFormat.Xlsm);
                Console.WriteLine("Workbook saved as 'VbaProjectProtected.xlsm'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception caught while saving workbook: {ex.Message}");
            }
        }
    }
}
