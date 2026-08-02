// Title: C# – Protect Aspose.Cells Workbook Structure (Block New Sheets, Allow Renames)
// Description: Demonstrates how to apply structure protection to an Aspose.Cells workbook with a password, preventing the insertion or deletion of worksheets while still permitting the renaming of existing tabs. The example creates a workbook, adds a second sheet, protects the structure, renames the first sheet, and saves the file.
// Keywords: Aspose.Cells workbook structure protection | C# prevent adding worksheets | allow sheet rename after protection | structure password Aspose.Cells | protect workbook layout .NET
// Common Searches: Aspose.Cells block new worksheets C# | rename sheet after structure protection Aspose | how to lock workbook layout but allow renaming | C# protect workbook structure with password | prevent sheet insertion Aspose.Cells
// Developer Intent: Apply a password‑protected structure lock that stops users from adding or removing worksheets while still letting them rename the existing ones.
// Use Cases: Distribute a template where the sheet count is fixed but users can label each tab for their scenario. | Secure a financial model so the tab order cannot be altered, yet analysts may update sheet names for clarity. | Create a report that forbids extra worksheets but permits the author to rename sections before finalizing.
// AI Prompts: Generate C# code using Aspose.Cells to lock workbook structure with a password and still allow sheet renaming. | Explain the steps to enable structure protection in Aspose.Cells so new worksheets cannot be added or deleted. | Show how to remove or change the structure protection password in an Aspose.Cells workbook after it has been saved.

using System;
using Aspose.Cells;

namespace AsposeCellsWorkbookProtectionDemo
{
    // Demonstrates how to apply structure protection to an Aspose.Cells workbook with a password, preventing the insertion or deletion of worksheets while still permitting the renaming of existing tabs. The example creates a workbook, adds a second sheet, protects the structure, renames the first sheet, and saves the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a second worksheet for demonstration
            workbook.Worksheets.Add("SecondSheet");

            // Protect the workbook structure (prevents adding/removing worksheets,
            // but allows renaming existing worksheets) with a password
            workbook.Protect(ProtectionType.Structure, "pwd123");

            // Rename the first worksheet (allowed under structure protection)
            workbook.Worksheets[0].Name = "RenamedFirstSheet";

            // Save the protected workbook
            workbook.Save("ProtectedWorkbook.xlsx");
        }
    }
}
