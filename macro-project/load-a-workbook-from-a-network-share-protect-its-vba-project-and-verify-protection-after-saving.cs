// Title: How to load an XLSM workbook from a UNC network share, protect its VBA project with a password, and verify protection using Aspose.Cells for .NET
// Common Searches: C# example for protecting a VBA project in an XLSM file stored on a UNC share with Aspose.Cells | Load workbook from a network path and set VBA password using Aspose.Cells .NET API | Verify VBA project protection after saving a macro-enabled workbook in C# | Aspose.Cells protect VBA project on a remote file server | Save a password‑protected macro workbook to a UNC location in C#
// Tags: load XLSM from UNC path Aspose.Cells | protect VBA project password C# | save macro-enabled workbook with protection Aspose.Cells | check VBA IsProtected property C# | Aspose.Cells VBA project protection example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an XLSM workbook from a UNC network share, ensures the source file and destination directory exist, applies password protection to the workbook's VBA project via Aspose.Cells, saves the protected file back to a UNC location, reloads it, and confirms that the VBA project remains protected.
class ProtectVbaProject
{
    static void Main()
    {
        // Paths to the workbook on a network share (UNC paths)
        string sourcePath = @"\\ServerName\ShareFolder\MyWorkbook.xlsm";
        string destinationPath = @"\\ServerName\ShareFolder\MyWorkbook_Protected.xlsm";

        try
        {
            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Ensure the destination directory exists
            string destDir = Path.GetDirectoryName(destinationPath);
            if (!Directory.Exists(destDir))
                Directory.CreateDirectory(destDir);

            // Load the workbook from the network share
            Workbook workbook = new Workbook(sourcePath);

            // Ensure the workbook contains a VBA project
            if (workbook.VbaProject != null)
            {
                // Protect the VBA project with a password (first argument: isReadOnly = false)
                string vbaPassword = "MySecretPassword";
                workbook.VbaProject.Protect(false, vbaPassword);
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Save the workbook (preserving the VBA project protection)
            workbook.Save(destinationPath, SaveFormat.Xlsm);

            // Reload the saved workbook to verify protection
            Workbook verifyWorkbook = new Workbook(destinationPath);

            // Check if the VBA project is protected
            bool isProtected = verifyWorkbook.VbaProject != null && verifyWorkbook.VbaProject.IsProtected;

            Console.WriteLine(isProtected
                ? "VBA project is successfully protected."
                : "VBA project protection failed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
