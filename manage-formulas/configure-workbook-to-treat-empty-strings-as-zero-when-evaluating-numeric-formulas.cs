using System;
using Aspose.Cells;
using System.Data;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put an empty string in A1
        worksheet.Cells["A1"].PutValue("");

        // Set a numeric formula that references the empty string cell
        // The formula should treat the empty string as zero, resulting in 10
        worksheet.Cells["B1"].Formula = "=A1+10";

        // Initialize WorkbookDesigner and configure it to replace empty strings with null
        // Null values are evaluated as zero in numeric formulas
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            UpdateEmptyStringAsNull = true
        };

        // Process the workbook (no external data source is required for this example)
        designer.Process();

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Display the calculated result (expected output: 10)
        Console.WriteLine("Calculated value in B1: " + worksheet.Cells["B1"].Value);

        // Save the workbook to verify the result
        workbook.Save("EmptyStringAsZero.xlsx");
    }
}