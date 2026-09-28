// Title: Programmatically unlock a password‑protected VBA project in an XLSM workbook using Aspose.Cells for .NET and save the unprotected file
// AI Prompts: Write C# code that opens an .xlsm file with Aspose.Cells, detects whether its VBA project is protected, and attempts to unprotect it using a supplied password. | Demonstrate how to use reflection to call the VbaProject.Unprotect(string) method when the method is not directly exposed in the current Aspose.Cells API. | Add comprehensive error handling that distinguishes between an incorrect password, missing Unprotect method, and other runtime exceptions, and reports the unlock status.
// Common Searches: aspocells c# how to remove VBA password protection from an xlsm workbook | unlock protected macro project programmatically using Aspose.Cells .NET | detect if VBA project is locked in an Excel file with Aspose.Cells | save an xlsm file after removing VBA protection with C# | reflection based unprotect for VBA project in Aspose.Cells
// Tags: Aspose.Cells VBA project unprotect | C# unlock password protected XLSM | reflection invoke Unprotect method | programmatic removal of VBA password | save unprotected macro-enabled workbook

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example loads an XLSM workbook, checks for a protected VBA project, uses reflection to invoke the Unprotect(string) method with a provided password, reports whether the unlock succeeded, and saves the workbook without VBA protection when successful.
class Program
{
    static void Main()
    {
        // Path to the workbook that contains a VBA project locked for viewing
        string inputPath = "input.xlsm";

        // Password that should unlock the VBA project
        string password = "YourPasswordHere";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            bool unlocked = false;

            // Check if a VBA project exists and is protected
            if (workbook.VbaProject != null && workbook.VbaProject.IsProtected)
            {
                try
                {
                    // Use reflection to call Unprotect(string) if it exists in the current Aspose.Cells version
                    MethodInfo unprotectMethod = workbook.VbaProject.GetType()
                        .GetMethod("Unprotect", new[] { typeof(string) });

                    if (unprotectMethod != null)
                    {
                        unprotectMethod.Invoke(workbook.VbaProject, new object[] { password });
                        unlocked = !workbook.VbaProject.IsProtected;
                    }
                    else
                    {
                        Console.WriteLine("The Unprotect method is not available in this Aspose.Cells version.");
                    }
                }
                catch (TargetInvocationException tie)
                {
                    Console.WriteLine($"Error while unlocking VBA project: {tie.InnerException?.Message ?? tie.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while unlocking VBA project: {ex.Message}");
                }

                // Report the result
                Console.WriteLine(unlocked
                    ? "VBA project unlocked successfully."
                    : "Failed to unlock VBA project. Incorrect password or unsupported operation.");
            }
            else
            {
                Console.WriteLine("The workbook does not contain a protected VBA project.");
                unlocked = true; // No protection, treat as ready to save
            }

            // If unlocked (or there was no protection), optionally save the workbook without the lock
            if (unlocked && workbook.VbaProject != null && !workbook.VbaProject.IsProtected)
            {
                string outputPath = "output_unlocked.xlsm";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
