// Title: C# Aspose.Cells: Create a Whole‑Column Named Range and Apply VLOOKUP Across Worksheets
// Description: Demonstrates how to build a new workbook, define a named range that spans entire columns (Data!$B:$C), retrieve its address, add a second sheet, insert VLOOKUP formulas that reference the named range, calculate all formulas, display results, and save the file as an .xlsx document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells named range whole column | C# VLOOKUP named range | cross‑sheet formula Aspose.Cells | Aspose.Cells calculate formulas | define column range .NET workbook | Aspose.Cells example VLOOKUP | named range address Aspose.Cells
// Common Searches: how to create a whole column named range in Aspose.Cells C# | use named range with VLOOKUP across sheets Aspose.Cells | Aspose.Cells calculate VLOOKUP after setting named range | retrieve address of a named range Aspose.Cells | Aspose.Cells example for column‑wide named range
// Developer Intent: Define a column‑wide named range and reference it in VLOOKUP formulas on another worksheet.
// Use Cases: Create a named range "LookupTable" that points to Data!$B:$C for dynamic lookups. | Insert VLOOKUP formulas on a separate sheet that pull values from the named range. | Programmatically evaluate all formulas and persist the results by saving the workbook.
// AI Prompts: Show me C# code that creates a whole‑column named range in Aspose.Cells and uses it in a VLOOKUP on a different sheet. | Provide an Aspose.Cells example that calculates VLOOKUP formulas referencing a column‑wide named range and saves the workbook. | Explain how to get and display the address of a named range defined as an entire column using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeVLookup
{
    // Demonstrates how to build a new workbook, define a named range that spans entire columns (Data!$B:$C), retrieve its address, add a second sheet, insert VLOOKUP formulas that reference the named range, calculate all formulas, display results, and save the file as an .xlsx document using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sheet1 – source data for VLOOKUP
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Data";

                // Populate column B with lookup keys and column C with corresponding values
                sheet1.Cells["B1"].PutValue("Apple");
                sheet1.Cells["C1"].PutValue(100);
                sheet1.Cells["B2"].PutValue("Banana");
                sheet1.Cells["C2"].PutValue(200);
                sheet1.Cells["B3"].PutValue("Cherry");
                sheet1.Cells["C3"].PutValue(300);
                sheet1.Cells["B4"].PutValue("Date");
                sheet1.Cells["C4"].PutValue(400);
                sheet1.Cells["B5"].PutValue("Elderberry");
                sheet1.Cells["C5"].PutValue(500);

                // -------------------------------------------------
                // Create a named range that refers to the whole columns B:C
                // -------------------------------------------------
                int nameIndex = workbook.Worksheets.Names.Add("LookupTable");
                workbook.Worksheets.Names[nameIndex].RefersTo = "=Data!$B:$C";

                // Retrieve the Range object via GetRange()
                Name lookupName = workbook.Worksheets.Names[nameIndex];
                AsposeRange lookupRange = lookupName.GetRange(); // whole columns B:C
                Console.WriteLine($"Named range address: {lookupRange.Address}");

                // -------------------------------------------------
                // Sheet2 – where VLOOKUP will be used
                // -------------------------------------------------
                Worksheet sheet2 = workbook.Worksheets.Add("Lookup");
                sheet2.Cells["A1"].PutValue("Cherry");
                sheet2.Cells["A2"].PutValue("Apple");
                sheet2.Cells["A3"].PutValue("Elderberry");

                // Apply VLOOKUP formula using the named range.
                sheet2.Cells["B1"].Formula = "=VLOOKUP(A1,LookupTable,2,FALSE)";
                sheet2.Cells["B2"].Formula = "=VLOOKUP(A2,LookupTable,2,FALSE)";
                sheet2.Cells["B3"].Formula = "=VLOOKUP(A3,LookupTable,2,FALSE)";

                // Calculate all formulas so that the results are stored in the cells
                workbook.CalculateFormula();

                // Display the results in the console (optional verification)
                Console.WriteLine($"Result for {sheet2.Cells["A1"].StringValue}: {sheet2.Cells["B1"].Value}");
                Console.WriteLine($"Result for {sheet2.Cells["A2"].StringValue}: {sheet2.Cells["B2"].Value}");
                Console.WriteLine($"Result for {sheet2.Cells["A3"].StringValue}: {sheet2.Cells["B3"].Value}");

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "NamedRangeVLookupDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
