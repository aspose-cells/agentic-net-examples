// Title: How to Enable AutoFilter on a Worksheet Header Row with Aspose.Cells for .NET (C#)
// Description: C# sample that creates a workbook, populates cells A1:C4 with product data, applies an AutoFilter to the header row (A1:C4), sets HasHeaders=true for proper column sorting, and saves the file as AutoFilterEnabled.xlsx.
// Keywords: Aspose.Cells | AutoFilter | C# | .NET | worksheet filter | header row | column sorting | Excel automation | filter arrows | Set AutoFilter range | Sorter.HasHeaders
// Common Searches: Aspose.Cells enable AutoFilter C# | Set AutoFilter range Aspose.Cells .NET | AutoFilter header row Aspose.Cells example | How to sort columns with AutoFilter using Aspose.Cells | C# code to add filter arrows to Excel sheet | Aspose.Cells AutoFilter sample
// Developer Intent: Add an AutoFilter to the first row of a worksheet so users can filter and sort columns directly in the generated Excel file.
// Use Cases: Generate a product catalog workbook where users can quickly filter by category or price. | Create export files for business analysts that include clickable filter arrows for ad‑hoc data analysis. | Prepare reporting sheets that require sortable headers without manual Excel configuration.
// AI Prompts: Show me C# code to enable AutoFilter on a specific range and set HasHeaders=true using Aspose.Cells. | How can I programmatically sort a column after applying AutoFilter with Aspose.Cells for .NET? | Give an example of determining the used range dynamically and applying AutoFilter with headers in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// C# sample that creates a workbook, populates cells A1:C4 with product data, applies an AutoFilter to the header row (A1:C4), sets HasHeaders=true for proper column sorting, and saves the file as AutoFilterEnabled.xlsx.
class EnableAutoFilterDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with a header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue("Electronics");
            sheet.Cells["C2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Shirt");
            sheet.Cells["B3"].PutValue("Clothing");
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["A4"].PutValue("Phone");
            sheet.Cells["B4"].PutValue("Electronics");
            sheet.Cells["C4"].PutValue(800);

            // Enable AutoFilter on the header row (first row) covering the data range
            sheet.AutoFilter.Range = "A1:C4";

            // Specify that the range contains headers so sorting works correctly
            sheet.AutoFilter.Sorter.HasHeaders = true;

            // Save the workbook
            workbook.Save("AutoFilterEnabled.xlsx");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        EnableAutoFilterDemo.Run();
    }
}
