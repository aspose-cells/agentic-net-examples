// Title: How to enable AutoFilter on a worksheet header row with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, adds a header row, applies an AutoFilter to the header range, and saves the file using Aspose.Cells. | Adjust the example to automatically find the last used column and set the AutoFilter range to cover the entire header row in a .NET workbook. | Show how to activate column sorting after the AutoFilter has been applied to a worksheet with Aspose.Cells.
// Common Searches: Aspose.Cells C# enable autofilter on first row of Excel worksheet | set auto filter range dynamically based on data columns using Aspose.Cells .NET | how to allow column sorting with auto filter in Aspose.Cells generated Excel file
// Tags: Aspose.Cells apply AutoFilter to header | C# set AutoFilter range dynamically | Excel column sorting with Aspose.Cells | save workbook with AutoFilter using Aspose.Cells | detect last column for AutoFilter Aspose.Cells

using Aspose.Cells;

// // This example creates a new workbook, writes a header row (A1:D1), enables AutoFilter on that range, and saves the workbook as AutoFilterDemo.xlsx.
class AutoFilterExample
{
    static void Main()
    {
        // Create a new workbook (using the create rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Optionally add some sample data with a header row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Category");
        sheet.Cells["D1"].PutValue("Price");

        // Enable AutoFilter on the header row (A1:D1)
        sheet.AutoFilter.Range = "A1:D1";

        // Save the workbook (using the save rule)
        workbook.Save("AutoFilterDemo.xlsx");
    }
}
