// Title: Get the Address of a Dynamic Named Range (SalesData) with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, defines a dynamic named range "SalesData" using an OFFSET formula, retrieves the corresponding Range object via GetRange(), and writes the range's address to the console. The workbook is then saved as an optional step.
// Keywords: Aspose.Cells | C# | .NET | dynamic named range | OFFSET formula | GetRange | named range address | Excel automation | retrieve range address | workbook saving
// Common Searches: Aspose.Cells get address of dynamic named range | C# retrieve named range address using Aspose.Cells | How to use OFFSET to define a named range in Aspose.Cells | GetRange example Aspose.Cells .NET | Log named range address in C# Excel library
// Developer Intent: Retrieve the address of the dynamic named range "SalesData" and output it (or use it programmatically).
// Use Cases: Confirm that a dynamically sized range includes the expected rows before running calculations. | Pass the range address to formulas, external reports, or APIs that require an A1‑style reference. | Log the address for debugging when generating Excel files automatically.
// AI Prompts: Show how to return the dynamic range address as a string instead of printing it. | Provide code to iterate over each cell in the retrieved SalesData range and read its values. | Explain how to handle the case where the named range "SalesData" is missing or has an invalid formula.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, defines a dynamic named range "SalesData" using an OFFSET formula, retrieves the corresponding Range object via GetRange(), and writes the range's address to the console. The workbook is then saved as an optional step.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                RetrieveDynamicNamedRangeAddress.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class RetrieveDynamicNamedRangeAddress
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate sample data in column A
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                }

                // Define a dynamic named range "SalesData"
                int nameIndex = workbook.Worksheets.Names.Add("SalesData");
                Name salesName = workbook.Worksheets.Names[nameIndex];
                salesName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

                // Retrieve the range that the name refers to
                Aspose.Cells.Range salesRange = salesName.GetRange();

                // Log the address of the dynamic named range
                Console.WriteLine($"Dynamic Named Range \"SalesData\" Address: {salesRange.Address}");

                // Save the workbook (optional, just to demonstrate lifecycle)
                workbook.Save("DynamicNamedRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
                throw;
            }
        }
    }
}
