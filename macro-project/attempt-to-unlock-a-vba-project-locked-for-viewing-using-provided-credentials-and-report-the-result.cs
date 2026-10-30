// Title: Unlock a password‑protected VBA project in an XLSM workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unlock VBA project password in .xlsm file | programmatically remove VBA project protection using Aspose.Cells .NET | how to use Aspose.Cells VbaProject.Unprotect method via reflection | check if Excel workbook VBA project is locked and unprotect it with C# | load .xlsm workbook and unprotect VBA macros using Aspose.Cells
// Tags: Aspose.Cells VBA project unprotect C# | unlock VBA macros in XLSM using Aspose.Cells | check VbaProject.IsProtected Aspose.Cells | invoke Unprotect via reflection Aspose.Cells | load workbook with VBA project Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VbaProject members

namespace VbaUnlockExample
{
    // The example loads an XLSM workbook, detects whether it contains a VBA project, checks if the project is protected, and if so attempts to unlock it by invoking the Unprotect method via reflection with the supplied password, then reports the success or failure of the operation.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file that contains a VBA project locked for viewing
            string inputFile = "LockedWorkbook.xlsm";

            // Password to unlock the VBA project (replace with actual password)
            string vbaPassword = "YourPassword";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Check if the workbook contains a VBA project
            if (workbook.VbaProject != null)
            {
                // Determine if the VBA project is protected (locked)
                if (workbook.VbaProject.IsProtected)
                {
                    try
                    {
                        // Attempt to unprotect the VBA project using reflection.
                        // This avoids compile‑time dependency on a specific API version.
                        var vbaProject = workbook.VbaProject;
                        var unprotectMethod = vbaProject.GetType().GetMethod("Unprotect", new[] { typeof(string) });

                        if (unprotectMethod != null)
                        {
                            unprotectMethod.Invoke(vbaProject, new object[] { vbaPassword });

                            // Verify the unprotect result
                            if (!vbaProject.IsProtected)
                            {
                                Console.WriteLine("VBA project successfully unlocked.");
                            }
                            else
                            {
                                Console.WriteLine("Failed to unlock VBA project. The password may be incorrect.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Unprotect method is not available in the current Aspose.Cells version.");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle any exceptions thrown by the unprotect operation
                        Console.WriteLine($"An error occurred while attempting to unlock the VBA project: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("VBA project is not locked for viewing.");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }

            // Optionally, save the workbook after unlocking (uncomment if needed)
            // string outputFile = "UnlockedWorkbook.xlsm";
            // try
            // {
            //     workbook.Save(outputFile);
            //     Console.WriteLine($"Workbook saved as \"{outputFile}\".");
            // }
            // catch (Exception ex)
            // {
            //     Console.WriteLine($"Failed to save workbook: {ex.Message}");
            // }
        }
    }
}
