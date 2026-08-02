// Title: C# – Protect Excel Workbook Structure While Allowing Cell Editing with Aspose.Cells
// Description: Demonstrates how to create a workbook, add worksheets, and apply password‑protected structure protection (ProtectionType.Structure) using Aspose.Cells for .NET. The protection blocks adding, deleting, renaming, or moving sheets, yet all cells remain editable because individual worksheets are left unprotected. The file is saved as ProtectedWorkbookStructure.xlsx.
// Keywords: Aspose.Cells | C# workbook protection | Excel structure protection .NET | ProtectionType.Structure | password protected workbook | allow cell editing | prevent sheet addition deletion | workbook.Protect example | Excel security C# | Aspose.Cells tutorial
// Common Searches: protect workbook structure Aspose.Cells C# | Aspose.Cells prevent sheet deletion while editing cells | how to lock Excel sheet order with password in .NET | Aspose.Cells structure only protection example | C# code to protect workbook layout but allow data entry
// Developer Intent: Apply password‑protected structure protection to an Excel workbook while keeping all cells editable.
// Use Cases: Distribute a template where users can enter data but cannot rearrange or remove worksheets. | Publish a financial report that must retain its original sheet order and count. | Automate an export process that safeguards workbook layout against accidental changes.
// AI Prompts: Write C# code using Aspose.Cells to protect only the workbook structure with a password and save the file. | Show how to enable ProtectionType.Structure in Aspose.Cells while leaving worksheets unprotected for editing. | Explain the difference between workbook.Protect and worksheet.Protect in Aspose.Cells and when to use each.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add worksheets, and apply password‑protected structure protection (ProtectionType.Structure) using Aspose.Cells for .NET. The protection blocks adding, deleting, renaming, or moving sheets, yet all cells remain editable because individual worksheets are left unprotected. The file is saved as ProtectedWorkbookStructure.xlsx.
    public class ProtectWorkbookStructureDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a couple of worksheets (optional, just for demonstration)
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");

            // Protect only the workbook structure with a password.
            // This prevents adding, deleting, renaming, or moving worksheets,
            // while still allowing users to edit cell contents on any sheet
            // because the individual worksheets are not protected.
            workbook.Protect(ProtectionType.Structure, "myPassword123");

            // Save the workbook to a file
            string outputPath = "ProtectedWorkbookStructure.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
