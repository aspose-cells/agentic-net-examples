// Title: Create a dynamic named range using an INDEX formula that reads its row index from another cell with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to add a workbook name whose RefersTo property is an INDEX formula referencing a data range and an index cell, then assign that name to a worksheet cell. | Demonstrate how to recalculate the workbook after defining the dynamic named range and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# define named range with INDEX formula using a cell as the row number | Create a dynamic Excel named range based on another cell's value with Aspose.Cells for .NET | Set RefersTo property to an INDEX formula in Aspose.Cells and evaluate the result | Recalculate formulas after adding a named range in Aspose.Cells C# example
// Tags: Aspose.Cells define named range with INDEX function | C# set RefersTo property to formula | dynamic named range based on cell value Aspose.Cells | recalculate workbook formulas Aspose.Cells | save workbook as .xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, fills A1:A10 with numbers, puts an index value in B1, defines a named range "DynamicValue" whose RefersTo is an INDEX formula that selects from Data!A1:A10 based on Data!B1, assigns the name to cell C1, recalculates formulas so C1 shows the correct value, and saves the file as DynamicNamedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a friendly name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate A1:A10 with sample numeric data (10, 20, ..., 100)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue((i + 1) * 10);
            }

            // Cell B1 will contain the index (1‑based) that selects an item from A1:A10
            sheet.Cells["B1"].PutValue(3); // Example: selects the 3rd item (value 30)

            // Define a named range "DynamicValue" using an INDEX formula.
            // The formula returns the element from A1:A10 based on the index in B1.
            string indexFormula = "=INDEX(Data!A1:A10, Data!B1)";

            // Add the name to the workbook's name collection and set its reference
            int nameIndex = workbook.Worksheets.Names.Add("DynamicValue");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = indexFormula;

            // Demonstrate the named range by placing its result in C1
            sheet.Cells["C1"].Formula = "DynamicValue";

            // Recalculate formulas so C1 shows the correct value
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "DynamicNamedRange.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
