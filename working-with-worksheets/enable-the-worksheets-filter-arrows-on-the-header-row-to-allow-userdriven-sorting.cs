// Title: Show Filter Arrows & Enable Sorting on a Protected Worksheet with Aspose.Cells (C#)
// Description: Creates a workbook, adds a header row with sample data, applies an AutoFilter range so filter arrows appear, protects the sheet, and grants users permission to sort and filter while the worksheet remains locked.
// Keywords: Aspose.Cells C# filter arrows | AutoFilter range Aspose.Cells | protect worksheet allow sorting | Excel filter dropdown protected sheet | Aspose.Cells enable sorting on protected sheet
// Common Searches: Aspose.Cells show filter arrows on header | C# enable sorting on protected worksheet | AutoFilter with protection Aspose.Cells | how to keep filter arrows after protecting sheet | Aspose.Cells allow filtering on locked sheet
// Developer Intent: Display AutoFilter dropdown arrows on the header row and let end‑users sort or filter data even when the worksheet is protected.
// Use Cases: Distribute a sales report where the layout is locked but users can filter by product category. | Provide a template for price analysis that prevents editing but permits column sorting via filter arrows. | Create an inventory workbook that is read‑only for most cells yet still supports dynamic filtering for downstream users.
// AI Prompts: Generate C# code that adds filter arrows to a header row and protects the sheet while allowing sorting and filtering with Aspose.Cells. | Explain how to set an AutoFilter range and configure AllowSorting and AllowFiltering on a protected worksheet in Aspose.Cells for .NET. | Show step‑by‑step instructions to enable filter dropdowns on a protected Excel sheet using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, adds a header row with sample data, applies an AutoFilter range so filter arrows appear, protects the sheet, and grants users permission to sort and filter while the worksheet remains locked.
class EnableFilterArrows
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Category");
        sheet.Cells["C1"].PutValue("Price");

        // Add some sample data
        sheet.Cells["A2"].PutValue("Laptop");
        sheet.Cells["B2"].PutValue("Electronics");
        sheet.Cells["C2"].PutValue(1200);

        sheet.Cells["A3"].PutValue("Shirt");
        sheet.Cells["B3"].PutValue("Clothing");
        sheet.Cells["C3"].PutValue(45);

        sheet.Cells["A4"].PutValue("Phone");
        sheet.Cells["B4"].PutValue("Electronics");
        sheet.Cells["C4"].PutValue(800);

        // Apply AutoFilter to the range that includes the header row.
        // This makes the filter arrows appear on the header cells.
        sheet.AutoFilter.Range = "A1:C4";

        // Protect the worksheet but allow the user to sort and filter.
        sheet.Protect(ProtectionType.All);
        sheet.Protection.AllowSorting = true;
        sheet.Protection.AllowFiltering = true;

        // Save the workbook.
        workbook.Save("WorksheetWithFilterArrows.xlsx");
    }
}
