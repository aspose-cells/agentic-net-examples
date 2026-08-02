// Title: Create a Dynamic Named Range for Column O with Aspose.Cells for .NET (C#)
// Description: This example shows how to programmatically add a named range that automatically expands as new rows are added to column O. It uses an OFFSET formula, retrieves the range with GetRange(true) to reflect changes, and saves the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | dynamic named range | OFFSET formula | C# | .NET | GetRange recalculate | auto‑expand range | named range programmatically | Excel automation | column O range
// Common Searches: Aspose.Cells create dynamic named range | C# OFFSET named range column O | expand named range after adding rows Aspose.Cells | GetRange true recalculate named range | define named range that grows with data Aspose.Cells .NET
// Developer Intent: Programmatically define a named range that automatically expands when new rows are added to column O in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Calculate totals, averages, or other aggregates over a column that receives new entries without adjusting formulas. | Apply data validation or conditional formatting to a column whose size changes over time. | Link a chart series to a dynamic range so the chart updates as more data rows are appended. | Set a pivot‑table source range that adjusts automatically with incoming data. | Generate reports where the number of rows in a specific column varies per execution.
// AI Prompts: Show how to modify the OFFSET formula to start the dynamic range at O3 instead of O2. | Provide code to delete the DynamicO named range after it is no longer needed. | Explain how to bind a chart series to the DynamicO named range using Aspose.Cells. | Generate an example of using the dynamic named range inside an Excel formula. | Demonstrate updating the range after removing rows from column O.

using System;
using Aspose.Cells;

// This example shows how to programmatically add a named range that automatically expands as new rows are added to column O. It uses an OFFSET formula, retrieves the range with GetRange(true) to reflect changes, and saves the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate initial data in column O (column index 14)
            sheet.Cells["O1"].PutValue("Header");
            sheet.Cells["O2"].PutValue(10);
            sheet.Cells["O3"].PutValue(20);

            // Add a named range that expands automatically with data in column O
            int nameIndex = workbook.Worksheets.Names.Add("DynamicO");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$O$2,0,0,COUNTA({sheet.Name}!$O:$O)-1,1)";

            // Retrieve the range to verify the initial address
            Aspose.Cells.Range initialRange = dynamicName.GetRange();
            Console.WriteLine($"Initial dynamic range address: {initialRange.Address}");

            // Add more rows to column O
            sheet.Cells["O4"].PutValue(30);
            sheet.Cells["O5"].PutValue(40);

            // Recalculate the named range (GetRange with recalculate = true)
            Aspose.Cells.Range updatedRange = dynamicName.GetRange(true);
            Console.WriteLine($"Updated dynamic range address: {updatedRange.Address}");

            // Save the workbook
            string outputPath = "DynamicNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
