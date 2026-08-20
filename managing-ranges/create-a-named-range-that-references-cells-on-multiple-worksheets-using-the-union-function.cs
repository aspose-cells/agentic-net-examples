// Title: Create a Named UnionRange Across Multiple Worksheets with Aspose.Cells for .NET (C#)
// Description: This example shows how to build a workbook with two sheets, fill cells on Sheet1!A1:A3 and Sheet2!B1:B3, define a UnionRange using the address "Sheet1!A1:A3,Sheet2!B1:B3", assign the name "MyUnionRange", set a common value for all cells, and save the file as UnionRangeMultipleSheets.xlsx.
// Keywords: Aspose.Cells UnionRange C# | named range multiple sheets | CreateUnionRange example | union address Aspose.Cells | cross‑sheet range .NET | Aspose.Cells coding sample
// Common Searches: Aspose.Cells create UnionRange across worksheets | named range that spans multiple sheets C# | how to set a common value for UnionRange Aspose.Cells | UnionRange syntax for multi‑sheet address
// Developer Intent: The developer needs to define a UnionRange that includes cells from different worksheets, give it a name, and optionally assign a shared value using Aspose.Cells for .NET.
// Use Cases: Combine cells from several sheets into a single named range for consolidated formulas or reporting. | Apply identical data, formatting, or validation to non‑contiguous cells spread across worksheets. | Reference a cross‑sheet union range in charts, data validation lists, or pivot table sources.
// AI Prompts: Generate C# code with Aspose.Cells that creates a UnionRange covering three worksheets and sets a formula referencing the union. | Explain how to iterate through, read, and modify values of a UnionRange that spans multiple sheets in an existing workbook. | Show how to export a UnionRange to a CSV file while preserving sheet information using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example shows how to build a workbook with two sheets, fill cells on Sheet1!A1:A3 and Sheet2!B1:B3, define a UnionRange using the address "Sheet1!A1:A3,Sheet2!B1:B3", assign the name "MyUnionRange", set a common value for all cells, and save the file as UnionRangeMultipleSheets.xlsx.
    public class UnionRangeMultipleSheetsDemo
    {
        public static void Run()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
            sheet2.Name = "Sheet2";

            // Populate some data on Sheet1 (A1:A3)
            sheet1.Cells["A1"].PutValue("S1-1");
            sheet1.Cells["A2"].PutValue("S1-2");
            sheet1.Cells["A3"].PutValue("S1-3");

            // Populate some data on Sheet2 (B1:B3)
            sheet2.Cells["B1"].PutValue("S2-1");
            sheet2.Cells["B2"].PutValue("S2-2");
            sheet2.Cells["B3"].PutValue("S2-3");

            // Create a UnionRange that references ranges on both worksheets.
            // The address string can contain multiple ranges separated by commas,
            // each prefixed with its sheet name.
            string unionAddress = "Sheet1!A1:A3,Sheet2!B1:B3";

            // The sheetIndex parameter is the index of the sheet where the address is evaluated.
            // Using 0 (Sheet1) is sufficient for this scenario.
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange(unionAddress, 0);

            // Assign a name to the union range. The name can be used in formulas or elsewhere.
            unionRange.Name = "MyUnionRange";

            // Optionally set a common value for all cells in the union range.
            unionRange.Value = "Combined";

            // Save the workbook
            workbook.Save("UnionRangeMultipleSheets.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                UnionRangeMultipleSheetsDemo.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
