// Title: Aspose.Cells for .NET – Create a dynamic named range with OFFSET and COUNTA
// Description: This example builds a new workbook, fills column A with a header and items, then adds a named range called **DynamicList** whose RefersTo formula uses `OFFSET` together with `COUNTA` to automatically span all non‑empty cells below the header. The code also shows how to read the current address of the range and save the file.
// Keywords: Aspose.Cells C# dynamic named range | OFFSET function Excel | COUNTA formula | programmatic named range .NET | retrieve named range address | Excel dynamic list Aspose | C# Aspose.Cells example
// Common Searches: how to create a dynamic named range with offset in Aspose.Cells | c# aspose.cells named range that expands with new rows | using COUNTA with OFFSET for a dynamic list in .NET | retrieve address of a named range created programmatically | Aspose.Cells dynamic range for data validation
// Developer Intent: Add a named range that automatically adjusts its size as items are added or removed, using the OFFSET function combined with COUNTA.
// Use Cases: Populate a drop‑down list that grows when new entries are appended to the source column. | Link charts or formulas to a range that updates without manual re‑definition. | Validate the range size at runtime before applying further processing or exporting.
// AI Prompts: Generate C# code that modifies the OFFSET formula to skip blank cells within the column. | Show how to refresh the DynamicList named range after inserting rows at the top of the data table. | Provide an example of using the DynamicList named range as a data series source for a chart in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This example builds a new workbook, fills column A with a header and items, then adds a named range called **DynamicList** whose RefersTo formula uses `OFFSET` together with `COUNTA` to automatically span all non‑empty cells below the header. The code also shows how to read the current address of the range and save the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a friendly name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate column A with sample data that will form the dynamic list
            sheet.Cells["A1"].PutValue("Header");   // Header row (will be excluded from the list)
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");
            // Additional items can be added later; the OFFSET range will adjust automatically

            // Add a named range that references the dynamic list using the OFFSET function
            // OFFSET(start, rows, cols, height, width)
            // Start at A2 (first data row), no offset, height = number of non‑empty cells in column A minus the header
            int nameIdx = workbook.Worksheets.Names.Add("DynamicList");
            Name dynamicName = workbook.Worksheets.Names[nameIdx];
            dynamicName.RefersTo = "=OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,1)";

            // Demonstrate retrieving the range that the name currently refers to
            AsposeRange dynamicRange = dynamicName.GetRange();
            sheet.Cells["C1"].PutValue("Dynamic range address:");
            sheet.Cells["C2"].PutValue(dynamicRange.Address);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DynamicNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
