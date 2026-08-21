// Title: C# – Apply AutoFilter to Header Row and Freeze It with Aspose.Cells
// Description: Creates a new workbook, writes a header and sample data, sets an AutoFilter on range A1:C1, freezes the first row using FreezePanes at A2, and saves the file as AutoFilterWithFreeze.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells AutoFilter C# | freeze header row Aspose | FreezePanes Excel C# | apply autofilter Aspose.Cells | Excel filter dropdowns frozen header
// Common Searches: Aspose.Cells add autofilter to first row C# | how to freeze header row after applying autofilter Aspose | C# code to keep Excel filter dropdowns visible while scrolling | freeze panes on top row Aspose.Cells .NET | set autofilter range and freeze panes example
// Developer Intent: Add an AutoFilter to the worksheet’s header row and freeze that row so the filter controls remain visible during scrolling.
// Use Cases: Sales dashboards where users filter columns and need the header fixed. | Export templates that require filterable columns with a locked top row for large data sets. | Inventory spreadsheets that keep filter dropdowns accessible while scrolling through rows.
// AI Prompts: Generate C# code that applies an AutoFilter to a dynamic header range and freezes the header row with Aspose.Cells. | Show how to freeze the first row after setting an AutoFilter without altering column widths in Aspose.Cells for .NET. | Explain the steps to combine AutoFilter and FreezePanes so filter dropdowns stay visible in extensive worksheets.

using Aspose.Cells;

// Creates a new workbook, writes a header and sample data, sets an AutoFilter on range A1:C1, freezes the first row using FreezePanes at A2, and saves the file as AutoFilterWithFreeze.xlsx using Aspose.Cells for .NET.
class AutoFilterFreezeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate header row
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Category");
        sheet.Cells["C1"].PutValue("Price");

        // Populate some sample data
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue("Fruit");
        sheet.Cells["C2"].PutValue(1.20);

        sheet.Cells["A3"].PutValue("Carrot");
        sheet.Cells["B3"].PutValue("Vegetable");
        sheet.Cells["C3"].PutValue(0.80);

        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B4"].PutValue("Fruit");
        sheet.Cells["C4"].PutValue(1.10);

        // Apply AutoFilter to the header row (covers columns A‑C)
        sheet.AutoFilter.Range = "A1:C1";

        // Freeze the header row so the filter dropdowns stay visible while scrolling
        // Freeze at cell A2, freezing 1 row (the header) and 0 columns
        sheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("AutoFilterWithFreeze.xlsx", SaveFormat.Xlsx);
    }
}
