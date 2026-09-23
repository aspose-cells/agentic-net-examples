// Title: Create a multi‑sheet named range using the UNION function in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that adds a global named range called MyUnionRange which combines Sheet1!A1:A5 and Sheet2!B1:B5 using the UNION formula and saves the workbook. | Demonstrate how to set the RefersTo property to a UNION expression to define a named range that spans multiple worksheets in an Aspose.Cells workbook. | Explain the steps to create two worksheets, populate them, and then create a named range covering cells from both sheets using the UNION function in Aspose.Cells C#.
// Common Searches: asp.net aspocells create named range across multiple worksheets using UNION formula | how to define a global named range that references cells from different sheets in Aspose.Cells C# | Aspose.Cells example of UNION function in a named range | C# code to add a multi‑sheet named range with UNION and save as .xlsx using Aspose.Cells | using RefersTo property to set UNION range in Aspose.Cells workbook
// Tags: Aspose.Cells multi-sheet named range | UNION formula RefersTo property | global names collection .NET | C# define named range spanning worksheets | Excel workbook save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook with two worksheets, fills sample data, defines a global named range 'MyUnionRange' that uses the UNION formula to reference Sheet1!A1:A5 and Sheet2!B1:B5, and saves the file as NamedRangeUnion.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it to "Sheet1"
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet and rename it to "Sheet2"
            int sheet2Index = workbook.Worksheets.Add(); // Add returns the index of the new sheet
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "Sheet2";

            // Populate some sample data
            for (int i = 0; i < 5; i++)
            {
                sheet1.Cells[i, 0].PutValue($"S1_R{i + 1}");
                sheet2.Cells[i, 1].PutValue($"S2_R{i + 1}");
            }

            // Create a named range that references cells on multiple worksheets using UNION
            // The leading '=' is required for the formula string.
            string unionFormula = "=UNION(Sheet1!$A$1:$A$5,Sheet2!$B$1:$B$5)";

            // Add the named range to the workbook's global name collection
            int nameIndex = workbook.Worksheets.Names.Add("MyUnionRange");
            Name unionName = workbook.Worksheets.Names[nameIndex];
            unionName.RefersTo = unionFormula;

            // Define output file path
            string outputPath = "NamedRangeUnion.xlsx";

            // Ensure the directory exists
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
