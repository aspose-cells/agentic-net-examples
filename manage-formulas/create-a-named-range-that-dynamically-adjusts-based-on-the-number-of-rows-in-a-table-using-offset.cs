// Title: Aspose.Cells .NET: C# code to add a dynamic named range using OFFSET and COUNTA
// Description: This C# example creates a workbook, fills column A with sample data, and defines a named range called DynamicRange whose reference uses OFFSET together with COUNTA. The range automatically expands to cover all non‑empty rows, the formulas are calculated, and the file is saved as DynamicNamedRange.xlsx.
// Keywords: Aspose.Cells | C# dynamic named range | OFFSET function | COUNTA | .NET Excel API | programmatic named range | expand range with data | Aspose.Cells example | Excel dynamic range C# | named range formula
// Common Searches: Aspose.Cells create dynamic named range C# | OFFSET COUNTA named range Aspose.Cells .NET | how to make a range expand with new rows using Aspose.Cells | C# code for dynamic Excel range with OFFSET | Aspose.Cells example for auto‑sizing named range
// Developer Intent: Define a named range that automatically grows to include every populated row in a column.
// Use Cases: Supply a continuously updating source list for data‑validation dropdowns. | Drive chart series so the visual updates when rows are added or removed. | Use in aggregate formulas (SUM, AVERAGE, etc.) that must reflect the current data set.
// AI Prompts: Generate C# code with Aspose.Cells that creates a named range using OFFSET and COUNTA to adjust its height based on column A values. | Explain how to obtain the address and row count of a dynamic named range after calling CalculateFormula in Aspose.Cells. | Show how to modify the OFFSET formula to cap the dynamic range at 500 rows while still using COUNTA.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicNamedRange
{
    // This C# example creates a workbook, fills column A with sample data, and defines a named range called DynamicRange whose reference uses OFFSET together with COUNTA. The range automatically expands to cover all non‑empty rows, the formulas are calculated, and the file is saved as DynamicNamedRange.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in column A starting from A2 (A1 is a header)
                cells["A1"].PutValue("Item");
                for (int i = 2; i <= 10; i++) // 9 data rows as an example
                {
                    cells[$"A{i}"].PutValue($"Value {i - 1}");
                }

                // Add a named range that expands automatically with the number of rows in column A
                // The formula uses OFFSET with COUNTA to determine the height.
                // OFFSET(start, rows, cols, height, width)
                // start = Sheet1!$A$2 (first data cell)
                // rows = 0, cols = 0 (no offset)
                // height = COUNTA(Sheet1!$A$2:$A$1000) -> counts non‑empty cells in the column
                // width = 1 (single column)
                int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$A$2,0,0,COUNTA({sheet.Name}!$A$2:$A$1000),1)";

                // Ensure formulas are calculated and retrieve the evaluated range
                workbook.CalculateFormula();
                AsposeRange actualRange = dynamicName.GetRange();

                Console.WriteLine($"DynamicRange address: {actualRange.RefersTo}");
                Console.WriteLine($"Rows in range: {actualRange.RowCount}");

                // Save the workbook
                workbook.Save("DynamicNamedRange.xlsx");
                Console.WriteLine("Workbook saved as DynamicNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
