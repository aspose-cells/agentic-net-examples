// Title: C# – Aspose.Cells: Protect worksheet, allow formatting, block row insertion
// Description: Demonstrates how to create a workbook, enable cell/column/row formatting, disable row insertion, set a password, and apply full protection using Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet protection C# | allow formatting Aspose.Cells | disable row insertion Aspose.Cells | worksheet password protection .NET | custom worksheet permissions Aspose.Cells
// Common Searches: Aspose.Cells allow formatting but prevent row insertion | protect Excel sheet with password using Aspose.Cells C# | custom worksheet protection settings Aspose.Cells | how to block row insertion in Aspose.Cells worksheet | set worksheet protection options Aspose.Cells .NET
// Developer Intent: Apply worksheet protection that permits formatting while prohibiting row insertion, secured with a password.
// Use Cases: Distribute a template where users can style data but cannot alter row count. | Publish a report that allows visual adjustments without changing its structure. | Maintain data integrity in shared workbooks by enabling formatting rights and disabling row additions.
// AI Prompts: Write C# code with Aspose.Cells to protect a worksheet, enable formatting, and block row insertion using a password. | Explain how to modify Aspose.Cells protection to allow column insertion while keeping row insertion disabled. | Provide a step‑by‑step tutorial for applying custom worksheet protection settings and saving the file in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsAdvancedProtection
{
    // Demonstrates how to create a workbook, enable cell/column/row formatting, disable row insertion, set a password, and apply full protection using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Access the worksheet's protection settings
            Protection protection = sheet.Protection;

            // Allow formatting of cells, columns, and rows
            protection.AllowFormattingCell = true;
            protection.AllowFormattingColumn = true;
            protection.AllowFormattingRow = true;

            // Disallow insertion of rows (default is false, set explicitly for clarity)
            protection.AllowInsertingRow = false;

            // Optional: set a password for the protection
            protection.Password = "StrongPassword123";

            // Apply protection to the worksheet (protect all aspects)
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("AdvancedWorksheetProtection.xlsx");
        }
    }
}
