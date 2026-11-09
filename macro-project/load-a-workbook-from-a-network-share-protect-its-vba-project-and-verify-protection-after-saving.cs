// Title: Protect a VBA project in an .xlsm workbook stored on a UNC network share using Aspose.Cells for .NET and verify protection after saving
// Common Searches: C# Aspose.Cells protect VBA project in workbook located on network share | How to load .xlsm file from UNC path and set VBA password with Aspose.Cells | Verify VBA project protection after saving workbook using Aspose.Cells .NET | Save protected VBA project back to same UNC location with Aspose.Cells | Check if workbook contains VBA project before applying protection in C#
// Tags: apply VBA password Aspose.Cells .NET | load workbook from UNC network share C# | save workbook as Xlsm using Aspose.Cells | check VbaProject.IsProtected after saving | ensure VbaProject exists before applying password

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an .xlsm workbook from a UNC network share with Aspose.Cells for .NET, confirming the presence of a VBA project, applying a password-protected read‑only lock to the VBA project, saving the workbook back to the same share, reloading it, and verifying that the VBA project remains protected.
class ProtectVbaProjectExample
{
    static void Main()
    {
        // Path to the workbook on a network share (UNC path)
        string networkPath = @"\\ServerName\ShareFolder\SampleWorkbook.xlsm";

        // Verify that the workbook file exists before attempting to load it
        if (!File.Exists(networkPath))
        {
            Console.WriteLine($"File not found: {networkPath}");
            return;
        }

        try
        {
            // Load the workbook from the network location
            Workbook workbook = new Workbook(networkPath);

            // Ensure the workbook contains a VBA project before attempting protection
            if (workbook.VbaProject != null)
            {
                // Protect the VBA project with a password (first argument: isReadOnly flag)
                string vbaPassword = "MySecretPassword";
                workbook.VbaProject.Protect(true, vbaPassword);
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Save the workbook back to the network share (overwrites the original file)
            workbook.Save(networkPath, SaveFormat.Xlsm);

            // Reload the workbook to verify that the VBA project is protected
            Workbook verificationWorkbook = new Workbook(networkPath);
            bool isProtected = verificationWorkbook.VbaProject != null && verificationWorkbook.VbaProject.IsProtected;

            // Output verification result
            Console.WriteLine(isProtected
                ? "VBA project is protected after saving."
                : "VBA project protection failed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
