// Title: Aspose.Cells for .NET – Protect a Worksheet, Allow Formatting, Block Cell Editing (C#)
// Description: Demonstrates how to lock cells, enable cell‑formatting, and protect a worksheet using Aspose.Cells. The example saves the workbook, attempts to change a locked cell (throws), formats the cell (succeeds), and prints the worksheet's IsProtected status.
// Keywords: Aspose.Cells worksheet protection C# | allow cell formatting Aspose.Cells | lock cells prevent editing .NET | Protection.AllowFormattingCell | Protection.AllowEditingContent | IsProtected property Aspose.Cells | Excel sheet security programmatic
// Common Searches: protect Excel sheet but still allow formatting with Aspose.Cells | how to block editing of locked cells in C# Aspose.Cells | verify worksheet protection after calling Protect() | Aspose.Cells AllowFormattingCell example
// Developer Intent: Create a protected worksheet where users can style cells but cannot modify the values of locked cells, and programmatically confirm that the protection works.
// Use Cases: Enforce data integrity while permitting visual styling in shared workbooks. | Automate validation that locked cells reject value changes after protection. | Check the IsProtected flag to ensure a sheet is secured before distribution.
// AI Prompts: Write C# code with Aspose.Cells that protects a worksheet, enables cell formatting, disables value editing, and shows how to test the protection. | Show how to catch the exception thrown when trying to modify a locked cell after worksheet protection in Aspose.Cells for .NET. | Explain the effect of Protection.AllowFormattingCell and Protection.AllowEditingContent on worksheet security.

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsProtectionDemo
{
    // Demonstrates how to lock cells, enable cell‑formatting, and protect a worksheet using Aspose.Cells. The example saves the workbook, attempts to change a locked cell (throws), formats the cell (succeeds), and prints the worksheet's IsProtected status.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put initial value into cell A1
            Cell cell = cells["A1"];
            cell.PutValue("Original Value");

            // Ensure the cell is locked (default is locked, but set explicitly)
            Style style = cell.GetStyle();
            style.IsLocked = true;
            cell.SetStyle(style);

            // Access protection settings
            Protection protection = sheet.Protection;

            // Allow users to format cells
            protection.AllowFormattingCell = true;

            // Disallow editing contents of locked cells
            protection.AllowEditingContent = false;

            // Protect the worksheet (no password needed for this demo)
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("ProtectedWorksheet.xlsx");

            // ----- Verification -----

            // Attempt to change the cell value (should be blocked)
            try
            {
                cell.PutValue("New Value");
                Console.WriteLine("Cell value changed (protection not enforced).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to modify cell value as expected: " + ex.Message);
            }

            // Attempt to format the cell (should be allowed)
            try
            {
                Style fmtStyle = cell.GetStyle();
                fmtStyle.ForegroundColor = Color.Yellow;
                fmtStyle.Pattern = BackgroundType.Solid;
                cell.SetStyle(fmtStyle);
                Console.WriteLine("Cell formatting applied successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to format cell: " + ex.Message);
            }

            // Verify worksheet protection status
            Console.WriteLine("Worksheet IsProtected: " + sheet.IsProtected);
        }
    }
}
