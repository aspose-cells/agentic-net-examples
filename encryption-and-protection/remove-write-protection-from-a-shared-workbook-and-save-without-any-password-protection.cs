// Title: C# – Remove Write Protection from a Shared Excel Workbook with Aspose.Cells and Save Unprotected
// Description: Loads a write‑protected shared workbook, clears the WriteProtection password (or leaves it empty), optionally unprotects the shared workbook, and saves a new copy without any write protection using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# remove write protection | unprotect shared Excel workbook .NET | clear workbook password Aspose.Cells | save Excel file without protection | Workbook.Settings.WriteProtection | Aspose.Cells API example | C# Excel protection removal
// Common Searches: how to remove write protection from an Excel file using Aspose.Cells | Aspose.Cells unprotect shared workbook programmatically | C# clear Excel workbook password Aspose | remove write protection and save workbook with Aspose.Cells | Aspose.Cells example for disabling write protection
// Developer Intent: Strip all write protection from a shared Excel workbook and generate an unprotected version using Aspose.Cells for .NET.
// Use Cases: Automate the de‑protection of shared workbooks before bulk processing or publishing. | Create an unprotected copy of a password‑protected Excel file for downstream systems that cannot handle protection. | Validate and handle missing input files gracefully while removing workbook protection.
// AI Prompts: Generate C# code with Aspose.Cells that opens a write‑protected shared workbook, removes its write protection, and saves an unprotected copy. | Explain how Workbook.Settings.WriteProtection can be cleared to disable write protection in an Excel file using Aspose.Cells. | Provide robust error‑handling for a routine that checks file existence, removes write protection, and saves the workbook, including logging of exceptions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a write‑protected shared workbook, clears the WriteProtection password (or leaves it empty), optionally unprotects the shared workbook, and saves a new copy without any write protection using Aspose.Cells for .NET.
    class RemoveWriteProtection
    {
        public static void Run()
        {
            // Input workbook path (write‑protected shared workbook)
            string inputPath = "SharedProtectedWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Remove write protection if it is enabled
                if (workbook.Settings.WriteProtection.IsWriteProtected)
                {
                    // Setting an empty password clears the protection
                    workbook.Settings.WriteProtection.Password = string.Empty;
                }

                // If the workbook is also a shared workbook and has a password,
                // uncomment and provide the password to unprotect it:
                // workbook.UnprotectSharedWorkbook("sharedPassword");

                // Output workbook path (unprotected version)
                string outputPath = "UnprotectedWorkbook.xlsx";

                // Save the workbook without write protection
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            RemoveWriteProtection.Run();
        }
    }
}
