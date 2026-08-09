// Title: C# – Protect an Aspose.Cells Worksheet: allow selection, block editing (password protected)
// Description: Demonstrates how to use Aspose.Cells for .NET to protect a worksheet, enable selection of locked and unlocked cells, disable content editing, and apply a password before saving the file.
// Keywords: Aspose.Cells worksheet protection C# | allow cell selection Aspose.Cells | disable editing Aspose.Cells | worksheet password protection .NET | read‑only Excel workbook Aspose
// Common Searches: Aspose.Cells protect worksheet allow selection only | C# code to lock cells but still let users select them | how to set password on Aspose.Cells worksheet | prevent editing in Aspose.Cells while keeping cells selectable | read‑only Excel file with Aspose.Cells .NET
// Developer Intent: Protect a worksheet so users can select any cell but cannot modify its contents.
// Use Cases: Distribute a read‑only report where viewers can copy data without changing formulas. | Provide a template that lets users fill only predefined input cells while the rest of the sheet remains locked. | Share financial statements with external partners, allowing review of data but preventing any edits.
// AI Prompts: Generate C# code using Aspose.Cells to protect a worksheet, enable selection of locked and unlocked cells, disable content editing, and set a password. | Show an example that configures worksheet.Protection.AllowEditingContent = false, sets AllowSelectingLockedCell and AllowSelectingUnlockedCell to true, then calls worksheet.Protect(ProtectionType.All). | Explain how to protect a worksheet with Aspose.Cells while allowing specific unlocked cells to be edited.

using System;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    // Demonstrates how to use Aspose.Cells for .NET to protect a worksheet, enable selection of locked and unlocked cells, disable content editing, and apply a password before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the protection settings of the worksheet
            Protection protection = worksheet.Protection;

            // Prevent editing of locked cells
            protection.AllowEditingContent = false;

            // Allow users to select both locked and unlocked cells
            protection.AllowSelectingLockedCell = true;
            protection.AllowSelectingUnlockedCell = true;

            // Set a password for the protection (optional but recommended)
            protection.Password = "myPassword123";

            // Apply protection to the worksheet (protect all aspects)
            worksheet.Protect(ProtectionType.All);

            // Save the workbook to a file
            workbook.Save("ProtectedWorksheet.xlsx");
        }
    }
}
