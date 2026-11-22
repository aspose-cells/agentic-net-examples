// Title: Check that worksheet protection in Aspose.Cells for .NET keeps hidden rows and columns unchanged
// Common Searches: Aspose.Cells C# how to ensure hidden rows stay hidden after protecting a worksheet | verify column visibility after calling Worksheet.Protect in Aspose.Cells .NET | does Protect(ProtectionType.All) affect IsHidden property of rows in Aspose.Cells | C# code sample to test hidden row and column status before and after sheet protection with Aspose.Cells | Aspose.Cells worksheet protection impact on hidden rows and columns
// Tags: worksheet protection hidden rows preservation Aspose.Cells | protect sheet without changing column hidden state C# | Aspose.Cells Protect method hidden status verification | ProtectionType.All hidden rows behavior Aspose.Cells | C# hidden row verification after worksheet protect

using Aspose.Cells;
using System;

// The example creates a workbook, hides row 2 and column B, prints their IsHidden values, applies full protection with ProtectionType.All, and prints the values again to demonstrate that worksheet protection does not modify the hidden status of rows or columns.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "TestSheet";

            // Hide a specific row (Row 2) and column (Column B) using zero‑based indexes
            sheet.Cells.Rows[1].IsHidden = true;      // Row index 1 => Row 2
            sheet.Cells.Columns[1].IsHidden = true;   // Column index 1 => Column B

            // Verify hidden status before protecting the sheet
            Console.WriteLine("Row 2 hidden before protect: " + sheet.Cells.Rows[1].IsHidden);
            Console.WriteLine("Column B hidden before protect: " + sheet.Cells.Columns[1].IsHidden);

            // Protect the worksheet (all protection types)
            sheet.Protect(ProtectionType.All);

            // Verify hidden status after protecting the sheet
            Console.WriteLine("Row 2 hidden after protect: " + sheet.Cells.Rows[1].IsHidden);
            Console.WriteLine("Column B hidden after protect: " + sheet.Cells.Columns[1].IsHidden);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
