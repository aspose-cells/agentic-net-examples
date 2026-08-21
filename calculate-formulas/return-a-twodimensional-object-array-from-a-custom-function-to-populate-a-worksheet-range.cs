// Title: Using SetArrayFormula to populate an Excel range from a 2‑D object[][] in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a 3×3 object[][] and writes it to cells A1:C3 using Cells.SetArrayFormula in Aspose.Cells. | Show how to assign pre‑calculated values to an array‑formula range with SetArrayFormula, including the required row and column parameters. | Demonstrate saving the workbook after filling the range with a nested object array in Aspose.Cells.
// Common Searches: aspnet setarrayformula example with object[][] | populate Excel cells from two dimensional array using Aspose.Cells C# | return custom function values as object[][] for Excel worksheet | how to set pre‑calculated values for an array formula in Aspose.Cells | fill range A1:C3 with mixed data types using Aspose.Cells SetArrayFormula
// Tags: setarrayformula with object[][] Aspose.Cells | populate worksheet range from 2d array C# | precalculated array formula values Aspose.Cells | custom function returning object[][] for Excel | Aspose.Cells fill cells using nested object array

using System;
using Aspose.Cells;

// The example creates a new workbook, defines a 3×3 object[][] containing numbers and strings, and uses Cells.SetArrayFormula to write those values into the range A1:C3. After optionally recalculating formulas, the workbook is saved as CustomFunctionArrayDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare a two‑dimensional object array that will populate the range
            // Row count = 3, column count = 3
            object[][] values = new object[][]
            {
                new object[] { 10, "Apple", 1.5 },
                new object[] { 20, "Banana", 2.5 },
                new object[] { 30, "Cherry", 3.5 }
            };

            // Set an array formula in cell A1.
            // The actual formula is irrelevant because we provide pre‑calculated values.
            cells["A1"].SetArrayFormula(
                "=SUM(A1:C1)",               // dummy formula
                rowNumber: values.Length,    // number of rows to fill
                columnNumber: values[0].Length, // number of columns to fill
                options: new FormulaParseOptions(),
                values: values);

            // Optional: calculate other formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "CustomFunctionArrayDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
