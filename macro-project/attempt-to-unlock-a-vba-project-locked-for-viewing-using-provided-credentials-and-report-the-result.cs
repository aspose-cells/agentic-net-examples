// Title: Unlock a password‑protected VBA project in an Excel .xlsm workbook with Aspose.Cells for .NET (C#)
// Description: Loads an .xlsm file using Aspose.Cells, detects a VBA project, reports its protection status, validates a supplied password, removes protection with VbaProject.Protect(false, null) when the password is correct, and saves the workbook as a new unlocked file.
// Keywords: Aspose.Cells VBA unlock | C# remove VBA password | Aspose.Cells VbaProject.ValidatePassword | unprotect Excel macro project | programmatic VBA project unprotect .NET | save unlocked xlsm Aspose | batch VBA password removal
// Common Searches: how to unlock a protected VBA project using Aspose.Cells C# | Aspose.Cells validate VBA password and unprotect workbook | C# code to remove VBA password from .xlsm file | unprotect Excel macro project programmatically .NET | Aspose.Cells VbaProject.Protect false example
// Developer Intent: The developer needs to programmatically unlock a password‑protected VBA project in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Validate a user‑provided password and unprotect the VBA project before saving a new copy of the workbook. | Check whether an uploaded .xlsm file contains a VBA project and report its protection state. | Automate processing of multiple macro‑enabled workbooks to remove VBA protection in bulk.
// AI Prompts: Generate C# code that uses Aspose.Cells to unlock a VBA project with a given password, including comprehensive error handling and console output. | Explain the purpose of VbaProject.Protect(false, null) and how it removes VBA project protection in Aspose.Cells. | Create a version of the demo that scans a directory for .xlsm files, attempts to unlock each VBA project, and logs success or failure to a CSV file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads an .xlsm file using Aspose.Cells, detects a VBA project, reports its protection status, validates a supplied password, removes protection with VbaProject.Protect(false, null) when the password is correct, and saves the workbook as a new unlocked file.
public class UnlockVbaProjectDemo
{
    public static void Run(string filePath, string password)
    {
        // Verify that the input file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook containing the VBA project
            Workbook workbook = new Workbook(filePath);
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Report current protection status
            Console.WriteLine($"VBA Project IsProtected: {vbaProject.IsProtected}");

            // Attempt to unlock only if it is protected
            if (vbaProject.IsProtected)
            {
                // Validate the supplied password
                bool isPasswordValid = vbaProject.ValidatePassword(password);
                Console.WriteLine($"Password validation result: {isPasswordValid}");

                if (isPasswordValid)
                {
                    // Unprotect the VBA project (set IsProtected = false)
                    vbaProject.Protect(false, null);
                    Console.WriteLine("VBA project unlocked successfully.");
                }
                else
                {
                    Console.WriteLine("Failed to unlock VBA project: invalid password.");
                }
            }
            else
            {
                Console.WriteLine("VBA project is not protected; no action needed.");
            }

            // Save the workbook (optional, demonstrates that the project is now unprotected)
            string outputPath = "unlocked_" + Path.GetFileName(filePath);
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    // Entry point required for compilation
    public static void Main(string[] args)
    {
        // args[0] = path to the .xlsm file
        // args[1] = password for the VBA project (optional)

        if (args.Length == 0)
        {
            Console.WriteLine("Usage: UnlockVbaProjectDemo <inputFilePath> [password]");
            return;
        }

        string inputFile = args[0];
        string password = args.Length > 1 ? args[1] : string.Empty;

        UnlockVbaProjectDemo.Run(inputFile, password);
    }
}
