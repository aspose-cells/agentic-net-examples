// Title: Protect Excel workbook metadata (author & title) with write protection using Aspose.Cells for .NET
// Description: Creates a workbook, sets BuiltInDocumentProperties (Author, Title), applies WriteProtection with a password, protection author and a read‑only recommendation, then saves the file as ProtectedProperties.xlsx.
// Keywords: Aspose.Cells workbook protection | Excel metadata lock C# | write protection password Aspose | prevent editing document properties | read‑only recommendation Aspose.Cells
// Common Searches: Aspose.Cells how to lock author and title | C# write protect Excel file properties | set password for Excel workbook using Aspose | make Excel workbook read‑only with Aspose.Cells | prevent metadata changes in Excel via code
// Developer Intent: Apply write protection so the workbook’s author and title fields cannot be modified after saving.
// Use Cases: Distribute confidential reports where original author information must stay intact. | Share templates that require the creator’s metadata to remain unchanged. | Enforce read‑only access for collaborative workbooks while preserving initial document properties.
// AI Prompts: Generate C# code to add write protection to an existing workbook and lock its built‑in document properties with Aspose.Cells. | Explain how to update or remove the write‑protection password on a protected Excel file using Aspose.Cells. | Show how to detect if a workbook has write protection enabled before allowing property edits.

using System;
using Aspose.Cells;

// Creates a workbook, sets BuiltInDocumentProperties (Author, Title), applies WriteProtection with a password, protection author and a read‑only recommendation, then saves the file as ProtectedProperties.xlsx.
class ProtectWorkbookProperties
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set default document properties (author, title)
        workbook.BuiltInDocumentProperties.Author = "Original Author";
        workbook.BuiltInDocumentProperties.Title = "Confidential Report";

        // Apply write protection to prevent changes to these properties
        workbook.Settings.WriteProtection.Password = "protect123";   // password required to modify the file
        workbook.Settings.WriteProtection.Author = "Protected Author"; // author of the protection
        workbook.Settings.WriteProtection.RecommendReadOnly = true; // suggest opening as read‑only

        // Save the protected workbook
        workbook.Save("ProtectedProperties.xlsx");
    }
}
