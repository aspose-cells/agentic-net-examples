// Title: Apply password protection to a VBA project in an XLSM workbook loaded from a UNC network share with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsm file from a UNC path, uses Aspose.Cells to set a read‑only password on its VbaProject, saves the workbook back to the network location, and checks the IsProtected flag after reloading. | Show how to verify VBA project protection after saving a macro‑enabled workbook by re‑instantiating the Workbook object and inspecting the VbaProject.IsProtected property.
// Common Searches: Aspose.Cells protect VBA project password UNC path C# example | load macro-enabled Excel file from network share and set VBA protection using .NET | verify VBA project IsProtected after saving with Aspose.Cells | C# code to protect VBA project in .xlsm stored on a file server | how to save protected VBA project back to network location with Aspose.Cells
// Tags: apply password to VBA project Aspose.Cells | load XLSM from UNC share C# | save macro-enabled workbook with protected VBA | inspect VbaProject.IsProtected after save | network file share handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an .xlsm workbook from a UNC network share, applies a read‑only password to its VBA project via Aspose.Cells, saves the file back to the share, reloads it, and confirms the VBA project remains protected.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook on a network share
            string sourcePath = @"\\server\share\folder\Sample.xlsm";

            // Path where the protected workbook will be saved
            string destPath = @"\\server\share\folder\Sample_Protected.xlsm";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook from the network location
            Workbook workbook = new Workbook(sourcePath);

            // Verify that the workbook contains a VBA project
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Protect the VBA project with a password (first argument: isReadOnly)
            string vbaPassword = "MySecretPassword";
            workbook.VbaProject.Protect(true, vbaPassword);

            // Ensure destination directory exists
            string destDir = Path.GetDirectoryName(destPath);
            if (!Directory.Exists(destDir))
                Directory.CreateDirectory(destDir);

            // Save the workbook (must be saved as a macro-enabled format)
            workbook.Save(destPath, SaveFormat.Xlsm);

            // Reload the saved workbook to verify that the VBA project is protected
            Workbook verificationWorkbook = new Workbook(destPath);
            bool isProtected = verificationWorkbook.VbaProject != null && verificationWorkbook.VbaProject.IsProtected;

            Console.WriteLine($"VBA project protected: {isProtected}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
