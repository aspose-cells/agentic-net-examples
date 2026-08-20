// Title: Get the Address of a Dynamic Named Range (SalesData) with Aspose.Cells for .NET
// Description: Creates a workbook, defines a dynamic named range "SalesData" using an OFFSET‑COUNTA formula, forces recalculation, retrieves the resolved range with GetRange(true), and logs its address to the console. The file can then be saved if needed.
// Keywords: Aspose.Cells dynamic named range address | C# GetRange true Aspose.Cells | OFFSET COUNTA named range .NET | retrieve named range address Aspose | log Excel range address C#
// Common Searches: how to get address of a dynamic named range in Aspose.Cells | Aspose.Cells GetRange true example | C# retrieve OFFSET defined name address | Aspose.Cells dynamic range formula calculation | log named range address in .NET
// Developer Intent: Obtain and display the cell address of the dynamic named range "SalesData".
// Use Cases: Verify that a dynamic range captures the correct rows before exporting data. | Debug Excel report generation by logging the resolved range address. | Use the address to apply formatting, borders, or additional calculations programmatically.
// AI Prompts: Write C# code with Aspose.Cells that creates a dynamic named range using OFFSET and then prints its address. | Explain how GetRange(true) resolves a named range defined by OFFSET‑COUNTA and how to handle an empty source column. | Show how to log a named range address and subsequently apply a border style to that range using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, defines a dynamic named range "SalesData" using an OFFSET‑COUNTA formula, forces recalculation, retrieves the resolved range with GetRange(true), and logs its address to the console. The file can then be saved if needed.
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

            // Populate some sample data in column A (this will be the source for the dynamic range)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item{i + 1}");
            }

            // Define a dynamic named range "SalesData" using OFFSET and COUNTA
            int nameIndex = workbook.Worksheets.Names.Add("SalesData");
            Name salesName = workbook.Worksheets.Names[nameIndex];
            salesName.RefersTo = $"=OFFSET({sheet.Name}!$A$1,0,0,COUNTA({sheet.Name}!$A:$A),1)";

            // Recalculate formulas so the dynamic range resolves
            workbook.CalculateFormula();

            // Retrieve the range that the name refers to and output its address
            Aspose.Cells.Range salesRange = salesName.GetRange(true);
            Console.WriteLine($"SalesData address: {salesRange.Address}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("DynamicNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
