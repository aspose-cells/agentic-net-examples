// Title: C# – Password‑protect a Shared Excel Workbook with Aspose.Cells
// Description: Load an existing Excel file, enable shared mode, apply a password using Workbook.ProtectSharedWorkbook, save the workbook, and verify protection with Settings.IsProtected—all with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# shared workbook protection | protect shared Excel file password .NET | Workbook.ProtectSharedWorkbook example | load and save protected workbook Aspose | Settings.IsProtected check | Excel encryption Aspose.Cells | shared mode password C# | Aspose.Cells sample code GitHub
// Common Searches: how to password protect a shared workbook using Aspose.Cells | Aspose.Cells protect shared mode C# example | set password for shared Excel file .NET | verify shared workbook protection Aspose | load shared workbook and apply password Aspose.Cells
// Developer Intent: Load a shared Excel workbook, add password protection to its shared mode, and save the protected file.
// Use Cases: Enforce a password before multiple users can edit a shared workbook. | Distribute a shared workbook to partners while preventing unauthorized changes. | Automate validation that a saved workbook retains shared‑mode protection after deployment.
// AI Prompts: Generate C# code that opens a shared Excel workbook with Aspose.Cells, sets a password for shared mode, and saves it. | Explain how Workbook.Settings.IsProtected indicates shared‑mode protection after saving. | Provide a step‑by‑step guide to test password protection on a shared workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsSharedWorkbookProtection
{
    // Load an existing Excel file, enable shared mode, apply a password using Workbook.ProtectSharedWorkbook, save the workbook, and verify protection with Settings.IsProtected—all with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing shared workbook
            string inputPath = "SharedWorkbook.xlsx";

            // Load the shared workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook is marked as shared (optional if already shared)
            workbook.Settings.Shared = true;

            // Apply password protection to the shared mode
            string password = "mySecretPassword";
            workbook.ProtectSharedWorkbook(password);

            // Save the protected workbook
            string outputPath = "ProtectedSharedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Optional: Verify protection status
            Workbook loaded = new Workbook(outputPath);
            Console.WriteLine("Workbook is protected (shared mode): " + loaded.Settings.IsProtected);
        }
    }
}
