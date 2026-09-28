// Title: Create a dynamic named range with OFFSET and COUNTA in Aspose.Cells for .NET and use it in a dependent formula
// AI Prompts: Generate C# code that adds a workbook‑level named range called DynamicList using an OFFSET‑COUNTA reference to column A in Aspose.Cells. | Insert a COUNTA formula that references the DynamicList named range, evaluate it programmatically, and retrieve the result. | Save the workbook as a .xlsx file after the calculation and write the count value to the console.
// Common Searches: Aspose.Cells .NET how to define a named range that expands with data using OFFSET | C# create a dynamic list range in Excel with Aspose.Cells and reference it in formulas | calculate number of entries in a dynamic named range using COUNTA in Aspose.Cells | using OFFSET and COUNTA together to build auto‑expanding ranges in Aspose.Cells | save workbook after adding a dynamic named range with Aspose.Cells C#
// Tags: dynamic named range OFFSET Aspose.Cells | COUNTA formula referencing named range .NET | auto expanding range based on column data C# | dependent calculation with named range Aspose.Cells | save workbook dynamic range Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, fills column A with sample items, defines a workbook‑level named range "DynamicList" using an OFFSET‑COUNTA formula that automatically expands as data grows, places a COUNTA formula referencing this range in cell C1, calculates the result, prints the count to the console, and saves the file as DynamicNamedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a meaningful name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate column A with sample data that will form the dynamic list
            string[] items = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
            for (int i = 0; i < items.Length; i++)
            {
                // Write values starting at cell A1
                sheet.Cells[i, 0].PutValue(items[i]);
            }

            // Define a named range "DynamicList" that expands automatically with the data in column A
            // The formula uses OFFSET together with COUNTA to create a dynamic reference:
            // =OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)
            int nameIndex = workbook.Worksheets.Names.Add("DynamicList");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = "=OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)";

            // Use the named range in a dependent calculation.
            // Example: count the number of items in the dynamic list.
            // Place the formula in cell C1.
            sheet.Cells["C1"].Formula = "=COUNTA(DynamicList)";

            // Calculate the formula with default calculation options
            sheet.Cells["C1"].Calculate(new CalculationOptions());

            // Output the result to the console (optional)
            Console.WriteLine("Number of items in DynamicList: " + sheet.Cells["C1"].StringValue);

            // Save the workbook to a file
            workbook.Save("DynamicNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
