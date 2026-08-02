// Title: Check worksheet read‑only protection and retrieve its type with Aspose.Cells for .NET
// Description: Creates a workbook, protects the first worksheet with ProtectionType.All and a password, saves and reloads the file, then uses IsProtected and Protection.IsProtectedWithPassword to report sheet protection status and password usage. Also reads workbook Settings.IsProtected and its ProtectionType.
// Keywords: Aspose.Cells | C# worksheet protection | read‑only sheet | ProtectionType.All | IsProtected property | password‑protected worksheet | .NET Excel security | workbook protection type
// Common Searches: Aspose.Cells detect worksheet protection | C# check if Excel sheet is read‑only with Aspose.Cells | Get worksheet protection type Aspose.Cells .NET | Determine password protection on a worksheet using Aspose.Cells | How to read workbook protection type in Aspose.Cells
// Developer Intent: Find out whether a worksheet is protected as read‑only and whether the protection relies on a password or another protection mode.
// Use Cases: Prevent editing when a sheet is marked read‑only in a data‑entry application. | Audit Excel files for compliance by logging sheet protection status and password usage. | Apply additional encryption only when the workbook or worksheet is not already password protected.
// AI Prompts: Generate C# code with Aspose.Cells that loads an existing workbook, checks if the first worksheet is protected, determines if a password is required, and prints the protection type. | Show an example that creates a workbook, protects a worksheet with read‑only settings, saves it, reloads it, and reports the IsProtected flag, password protection flag, and workbook protection type.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, protects the first worksheet with ProtectionType.All and a password, saves and reloads the file, then uses IsProtected and Protection.IsProtectedWithPassword to report sheet protection status and password usage. Also reads workbook Settings.IsProtected and its ProtectionType.
    public class WorksheetReadOnlyCheckDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Protect the worksheet (read‑only style) with a password
                sheet.Protect(ProtectionType.All, "readonlypwd", null);

                // Save the workbook to a temporary file
                string filePath = "ReadOnlyWorksheetDemo.xlsx";
                workbook.Save(filePath);

                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File '{filePath}' was not found.");
                    return;
                }

                // Load the saved workbook
                Workbook loadedWorkbook = new Workbook(filePath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                // Check if the worksheet is protected
                bool isProtected = loadedSheet.IsProtected;

                // Determine if the protection is password‑based
                bool isPasswordProtected = loadedSheet.Protection.IsProtectedWithPassword;

                // Report the protection status
                Console.WriteLine($"Worksheet protected: {isProtected}");
                Console.WriteLine($"Protected with password: {isPasswordProtected}");

                // If the workbook itself is protected, report its protection type
                if (loadedWorkbook.Settings.IsProtected)
                {
                    ProtectionType wbProtection = loadedWorkbook.Settings.ProtectionType;
                    Console.WriteLine($"Workbook protection type: {wbProtection}");
                }
                else
                {
                    Console.WriteLine("Workbook is not protected.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetReadOnlyCheckDemo.Run();
        }
    }
}
