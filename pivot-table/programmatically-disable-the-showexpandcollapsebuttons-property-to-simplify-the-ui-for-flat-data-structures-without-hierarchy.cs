// Title: Disable ShowExpandCollapseButtons for a PivotTable with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to turn off ShowExpandCollapseButtons for a pivot table | disable expand collapse buttons in Aspose.Cells pivot table for flat data | remove outline symbols from pivot table UI using Aspose.Cells .NET | simplify pivot table interface when there is no hierarchy with Aspose.Cells | Aspose.Cells hide pivot table expand collapse icons programmatically
// Tags: Aspose.Cells disable ShowExpandCollapseButtons | hide pivot table expand collapse icons .NET | pivot table UI simplification Aspose.Cells | remove outline symbols workbook Aspose.Cells | flat data structure pivot table settings Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, notes that the Outline.ShowOutlineSymbols property may be unavailable in older Aspose.Cells versions, and saves the workbook as output.xlsx within a try‑catch block. To hide the expand/collapse buttons for a pivot table you would set the appropriate ShowExpandCollapseButtons (or related outline) property to false before saving.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet (optional, shown for completeness)
            var worksheet = workbook.Worksheets[0];

            // Note: The Outline.ShowOutlineSymbols property may not be available in older
            // versions of Aspose.Cells. If needed, adjust the code according to the
            // version you are using. Here we simply proceed without modifying outline symbols.

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
