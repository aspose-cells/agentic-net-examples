// Title: Convert String Numeric Values to Numbers in the First Worksheet with Aspose.Cells for .NET
// Description: Shows how to create a workbook, fill the first worksheet with numeric strings, apply Cells.ConvertStringToNumericValue only to that sheet, leave all other worksheets untouched, and save the file as an XLSX document.
// Keywords: Aspose.Cells ConvertStringToNumericValue | C# convert string to numeric | first worksheet conversion | preserve other sheets | numeric string to number .NET | Aspose.Cells workbook example
// Common Searches: Aspose.Cells convert string to numeric in one sheet | ConvertStringToNumericValue first worksheet C# | How to keep other worksheets unchanged when converting strings Aspose.Cells | C# Aspose.Cells numeric string conversion example | Convert numeric strings only on first sheet Aspose.Cells
// Developer Intent: Transform all cells that contain numeric strings into true numeric values, but only on the workbook's first worksheet.
// Use Cases: Clean imported text data on the main sheet before performing calculations. | Prepare a financial report where only the primary data sheet requires numeric conversion. | Validate and convert numeric strings prior to generating charts or formulas on the first worksheet.
// AI Prompts: Provide C# code that uses Aspose.Cells to convert numeric strings to numbers only on the first worksheet, leaving other sheets unchanged. | Show an example of calling ConvertStringToNumericValue for a specific worksheet and then saving the workbook. | Explain how to confirm that conversion happened solely on the first sheet while other worksheets retain their original string values.

using System;
using Aspose.Cells;

namespace AsposeCellsStringToNumericDemo
{
    // Shows how to create a workbook, fill the first worksheet with numeric strings, apply Cells.ConvertStringToNumericValue only to that sheet, leave all other worksheets untouched, and save the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with the default single worksheet
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet firstSheet = workbook.Worksheets[0];
            Cells firstCells = firstSheet.Cells;

            // Populate the first worksheet with string values that can be converted to numbers
            firstCells["A1"].PutValue("123");          // numeric string
            firstCells["A2"].PutValue("45.67");        // decimal string
            firstCells["A3"].PutValue("NotANumber");   // non‑numeric string (should stay as string)

            // Add a second worksheet to demonstrate that it remains unchanged
            Worksheet secondSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            Cells secondCells = secondSheet.Cells;
            secondCells["A1"].PutValue("890");         // numeric string but we will NOT convert it
            secondCells["A2"].PutValue("Hello");       // regular string

            // Convert string values to numeric where possible ONLY in the first worksheet
            firstCells.ConvertStringToNumericValue();

            // Verify conversion (optional console output)
            Console.WriteLine("First sheet A1 (numeric): " + firstCells["A1"].DoubleValue);
            Console.WriteLine("First sheet A2 (numeric): " + firstCells["A2"].DoubleValue);
            Console.WriteLine("First sheet A3 (string): " + firstCells["A3"].StringValue);
            Console.WriteLine("Second sheet A1 (still string): " + secondCells["A1"].StringValue);
            Console.WriteLine("Second sheet A2 (still string): " + secondCells["A2"].StringValue);

            // Save the workbook to a file
            workbook.Save("StringToNumericResult.xlsx");
        }
    }
}
