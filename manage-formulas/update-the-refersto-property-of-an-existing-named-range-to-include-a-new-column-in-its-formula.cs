// Title: Extend a Named Range RefersTo to an Additional Column using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, define a named range (MyRange) that points to A1:B5 on the "Data" sheet, update its RefersTo property to include column C (A1:C5), retrieve the updated Range object, and save the workbook as UpdatedNamedRange.xlsx.
// Keywords: Aspose.Cells named range RefersTo | C# extend named range column | modify named range formula Aspose.Cells | update RefersTo property .NET | dynamic range expansion Aspose.Cells
// Common Searches: Aspose.Cells change RefersTo of named range C# | add column to existing named range Aspose.Cells | extend named range programmatically .NET | update named range address Aspose.Cells | how to modify RefersTo property in Aspose.Cells
// Developer Intent: Programmatically modify an existing named range's RefersTo formula to include a new column.
// Use Cases: Automatically expand a data range when new columns are added, keeping charts and formulas up‑to‑date. | Adjust named ranges after column insertions to preserve reference integrity in reports. | Create flexible templates where the range size adapts to varying data layouts before saving.
// AI Prompts: Write C# code with Aspose.Cells that receives a workbook, a named range name, and a target column letter, then updates the RefersTo to include that column. | Show how to fetch a named range, change its RefersTo from "=Data!$A$1:$B$5" to "=Data!$A$1:$C$5", and confirm the new address using Aspose.Cells for .NET. | Create a method that expands any named range by a given number of columns while keeping the original row boundaries.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a named range (MyRange) that points to A1:B5 on the "Data" sheet, update its RefersTo property to include column C (A1:C5), retrieve the updated Range object, and save the workbook as UpdatedNamedRange.xlsx.
    public class UpdateNamedRangeRefersTo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some sample data in columns A and B (rows 1-5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue($"A{i + 1}");
                sheet.Cells[i, 1].PutValue($"B{i + 1}");
            }

            // Add a named range that currently refers to columns A and B
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRange = workbook.Worksheets.Names[nameIndex];
            // Initial reference: columns A and B, rows 1 to 5
            myRange.RefersTo = "=Data!$A$1:$B$5";

            Console.WriteLine("Original RefersTo: " + myRange.RefersTo);

            // Extend the range to include column C as well (rows 1-5)
            string newRefersTo = "=Data!$A$1:$C$5";

            // Update the RefersTo property of the existing named range
            myRange.RefersTo = newRefersTo;

            Console.WriteLine("Updated RefersTo: " + myRange.RefersTo);

            // Retrieve the range object and print its address
            Aspose.Cells.Range updatedRange = myRange.GetRange();
            Console.WriteLine("Updated range address: " + updatedRange.Address);

            // Save the workbook
            string outputPath = "UpdatedNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
