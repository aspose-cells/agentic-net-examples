// Title: Define a dynamic named range with OFFSET in Aspose.Cells for .NET
// Description: This example creates a workbook, fills A1:A10, adds a named range called DynamicRange that uses OFFSET combined with COUNTA to automatically adjust its height, retrieves the range's address and size, and saves the file as DynamicNamedRange.xlsx.
// Keywords: Aspose.Cells OFFSET named range | dynamic range C# | COUNTA OFFSET formula .NET | add named range programmatically | Aspose.Cells GetRange | C# Excel dynamic range
// Common Searches: Aspose.Cells create dynamic named range | OFFSET formula with COUNTA in C# | retrieve address of named range Aspose.Cells | save workbook after adding named range | C# Excel dynamic range using OFFSET
// Developer Intent: Programmatically add a named range that expands based on the count of non‑empty cells in a column.
// Use Cases: Link a chart to a range that grows as new rows are added. | Apply data validation that automatically includes incoming entries. | Set a pivot‑table source to adapt to varying row counts without manual updates.
// AI Prompts: Show how to change the OFFSET formula to start at B2 and span two columns. | Provide code to delete the 'DynamicRange' name and verify its removal. | Explain how to reference the dynamic named range in a SUM formula using Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This example creates a workbook, fills A1:A10, adds a named range called DynamicRange that uses OFFSET combined with COUNTA to automatically adjust its height, retrieves the range's address and size, and saves the file as DynamicNamedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);
            }

            // Add a named range that uses an OFFSET formula.
            // The formula creates a dynamic range starting at A1,
            // with a height equal to the number of non‑empty cells in column A,
            // and a width of 1 column.
            int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = "=OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)";

            // Retrieve the actual range that the name refers to
            AsposeRange dynamicRange = dynamicName.GetRange();
            Console.WriteLine("Dynamic range address: " + dynamicRange.Address);
            Console.WriteLine("Rows: " + dynamicRange.RowCount + ", Columns: " + dynamicRange.ColumnCount);

            // Save the workbook
            string outputPath = "DynamicNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
