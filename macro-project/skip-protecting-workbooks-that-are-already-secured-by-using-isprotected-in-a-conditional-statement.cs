// Title: C# – Conditional Workbook Protection with Aspose.Cells for .NET
// Description: Loads an Excel file, checks Workbook.Settings.IsProtected, applies password protection only when the workbook is not already secured, and saves the result, preventing double‑protection of existing files.
// Keywords: Aspose.Cells | C# workbook protection | Excel password protection | Settings.IsProtected | conditional protect workbook | protect Excel structure | skip existing protection | Workbook.Protect | SaveFormat.Xlsx
// Common Searches: Aspose.Cells check if workbook is protected | C# protect Excel workbook only if not already protected | How to use Settings.IsProtected with Aspose.Cells | Skip protecting already secured workbook Aspose.Cells | Conditional workbook protection .NET
// Developer Intent: Apply password protection to an Excel workbook only when it lacks existing protection, avoiding redundant security actions.
// Use Cases: Batch‑process a folder of spreadsheets, adding a password only to files that are unprotected. | Integrate into a document‑management pipeline to enforce security on newly uploaded workbooks without overwriting existing protections. | Create a macro that validates protection status before saving, complying with policies that forbid re‑protecting already secured files.
// AI Prompts: Generate C# code using Aspose.Cells that iterates through all .xlsx files in a directory and protects each workbook with a password only if workbook.Settings.IsProtected is false. | Provide an example that logs a warning when a workbook is already protected and skips the Protect call in Aspose.Cells for .NET. | Write a reusable method that accepts a file path and password, checks the workbook's protection status, applies protection if needed, and returns a boolean indicating whether protection was applied.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // Loads an Excel file, checks Workbook.Settings.IsProtected, applies password protection only when the workbook is not already secured, and saves the result, preventing double‑protection of existing files.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Check if the workbook structure or window is already protected
            if (!workbook.Settings.IsProtected)
            {
                // Protect the workbook with a password if it is not already protected
                workbook.Protect(ProtectionType.Structure, "MySecretPassword");
            }

            // Save the workbook (protected only if it was not already protected)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
