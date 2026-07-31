// Title: Add comments to Excel named ranges using Aspose.Cells for .NET (C#)
// Description: This C# example creates a new workbook, defines three named ranges (ProductList, PriceList, HeaderRow) on a sheet named "Data", assigns purpose‑driven comments that capture business rules via the Name.Comment property, and saves the file as NamedRangesWithComments.xlsx. The comments are visible in Excel's Name Manager, helping downstream processes understand each range's intent.
// Keywords: Aspose.Cells | C# | .NET | Excel named range comment | Name.Comment property | add comment to named range | Excel automation | metadata for named ranges | business rule documentation | Excel Name Manager
// Common Searches: how to set a comment on a named range using Aspose.Cells C# | Aspose.Cells add description to named range | store business rules in Excel named ranges programmatically | C# example for Name.Comment in Aspose.Cells | Excel named range documentation with Aspose.Cells
// Developer Intent: Programmatically attach explanatory comments that describe the purpose and validation rules of each named range in an Excel workbook.
// Use Cases: Ensure report generators recognize that the ProductList range must never be empty. | Validate that every entry in the PriceList range is greater than zero. | Lock the HeaderRow range for export templates, indicating it should remain unchanged.
// AI Prompts: Generate C# code that loops through all named ranges in a workbook and adds a default comment when the Comment property is empty, using Aspose.Cells. | Explain how the Name.Comment value is stored inside the XLSX file and how Excel displays it in the Name Manager UI. | Show how to update an existing named range's comment to reflect a new business rule without modifying its RefersTo reference.

using System;
using Aspose.Cells;

// This C# example creates a new workbook, defines three named ranges (ProductList, PriceList, HeaderRow) on a sheet named "Data", assigns purpose‑driven comments that capture business rules via the Name.Comment property, and saves the file as NamedRangesWithComments.xlsx. The comments are visible in Excel's Name Manager, helping downstream processes understand each range's intent.
class AddCommentsToNamedRanges
{
    static void Main()
    {
        // Create a new workbook (Excel 2007+ format)
        Workbook workbook = new Workbook();

        // Access the first worksheet and give it a meaningful name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate some sample data that the named ranges will refer to
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["B2"].PutValue(1.2);
        sheet.Cells["B3"].PutValue(0.8);

        // ------------------------------------------------------------
        // Define a named range for the list of products
        // ------------------------------------------------------------
        int prodIndex = workbook.Worksheets.Names.Add("ProductList");
        Name prodName = workbook.Worksheets.Names[prodIndex];
        prodName.RefersTo = "=Data!$A$2:$A$3";
        // Comment describing purpose and business rule
        prodName.Comment = "List of products used in sales reports. Must not be empty.";

        // ------------------------------------------------------------
        // Define a named range for the corresponding prices
        // ------------------------------------------------------------
        int priceIndex = workbook.Worksheets.Names.Add("PriceList");
        Name priceName = workbook.Worksheets.Names[priceIndex];
        priceName.RefersTo = "=Data!$B$2:$B$3";
        // Comment describing purpose and business rule
        priceName.Comment = "Corresponding prices. Business rule: price must be greater than zero.";

        // ------------------------------------------------------------
        // Define a global named range for the header row (visible to all sheets)
        // ------------------------------------------------------------
        int headerIndex = workbook.Worksheets.Names.Add("HeaderRow");
        Name headerName = workbook.Worksheets.Names[headerIndex];
        headerName.RefersTo = "=Data!$A$1:$B$1";
        // Comment describing purpose and business rule
        headerName.Comment = "Header row for export templates. Should remain unchanged.";

        // Save the workbook with the named ranges and their comments
        workbook.Save("NamedRangesWithComments.xlsx");
    }
}
