// Title: C# Aspose.Cells: Create a Dynamic Named Range with OFFSET and COUNTA
// Description: Learn how to add a workbook, populate column A, define a named range called DynamicRange using OFFSET + COUNTA, retrieve its address and size, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# dynamic named range | OFFSET function Excel | COUNTA formula | programmatic named range | Excel automation .NET | chart data source range
// Common Searches: Aspose.Cells OFFSET dynamic range example | C# create named range with COUNTA | retrieve address of dynamic range Aspose | save workbook after adding named range | use dynamic range for chart source in .NET
// Developer Intent: Programmatically define a range that automatically expands to include every non‑empty cell in column A.
// Use Cases: Supply a self‑adjusting data source for charts that grow as new rows are added. | Reference the range in formulas, data validation, or pivot tables without manual updates. | Export workbooks that other systems can read, knowing the range always reflects current data.
// AI Prompts: Generate C# Aspose.Cells code to create a dynamic named range based on non‑blank cells in column B. | Show how to modify the OFFSET formula to start at A2 and ignore the header row. | Provide robust error‑handling patterns when calling GetRange() on a dynamic name.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace DynamicNamedRangeDemo
{
    // Learn how to add a workbook, populate column A, define a named range called DynamicRange using OFFSET + COUNTA, retrieve its address and size, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                sheet.Name = "Sheet1";

                // Sample data in column A
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue(10);
                sheet.Cells["A3"].PutValue(20);
                sheet.Cells["A4"].PutValue(30);
                // Row 5 left blank to demonstrate dynamic sizing

                // Add a dynamic named range using OFFSET and COUNTA
                int nameIndex = wb.Worksheets.Names.Add("DynamicRange");
                Name dynamicRange = wb.Worksheets.Names[nameIndex];
                dynamicRange.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

                // Resolve the range
                AsposeRange resolvedRange = dynamicRange.GetRange();
                Console.WriteLine($"Dynamic range address: {resolvedRange.Address}");
                Console.WriteLine($"Rows in range: {resolvedRange.RowCount}");
                Console.WriteLine($"Columns in range: {resolvedRange.ColumnCount}");

                // Save the workbook
                string outputPath = "DynamicNamedRangeDemo.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
