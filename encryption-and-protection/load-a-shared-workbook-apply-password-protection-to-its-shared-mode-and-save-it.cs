// Title: C# – Apply Password Protection to a Shared Workbook with Aspose.Cells
// Description: Load an existing Excel file, enable shared mode, set a password using ProtectSharedWorkbook, save the workbook, and confirm the protection status with Settings.IsProtected—all with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# shared workbook protection | ProtectSharedWorkbook | password protected Excel | shared workbook encryption .NET | Workbook.Settings.IsProtected | Aspose.Cells example | Excel password protection C# | save protected workbook | Aspose.Cells API
// Common Searches: Aspose.Cells protect shared workbook C# | How to set password for shared Excel file using Aspose.Cells | ProtectSharedWorkbook method example | Verify shared workbook protection Aspose.Cells | C# code to encrypt shared workbook
// Developer Intent: Load a shared Excel workbook, apply a password to its shared mode, and save the protected file.
// Use Cases: Secure a collaborative workbook before distribution to prevent unauthorized edits. | Enforce read‑only access for team members by password‑protecting the shared mode. | Programmatically verify that a saved workbook remains protected after deployment.
// AI Prompts: Generate C# code that loads a workbook, enables shared mode, applies a password with ProtectSharedWorkbook, and saves it. | Show how to handle exceptions when protecting a shared workbook with a password using Aspose.Cells. | Explain how to change or remove the password of an already protected shared workbook in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsSharedWorkbookProtection
{
    // Load an existing Excel file, enable shared mode, set a password using ProtectSharedWorkbook, save the workbook, and confirm the protection status with Settings.IsProtected—all with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the existing shared workbook
            string inputPath = "SharedWorkbook.xlsx";

            // Load the shared workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook is marked as shared (optional if already shared)
            workbook.Settings.Shared = true;

            // Apply password protection to the shared mode
            string password = "MySecretPassword";
            workbook.ProtectSharedWorkbook(password);

            // Save the protected workbook
            string outputPath = "ProtectedSharedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Optional: verify protection status
            Workbook loaded = new Workbook(outputPath);
            Console.WriteLine("Workbook is protected: " + loaded.Settings.IsProtected);
        }
    }
}
