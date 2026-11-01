// Title: How to automatically update a dynamic named range after inserting rows using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# insert rows and keep named range dynamic | C# update OFFSET formula for named range after adding rows in Excel workbook | Refresh dynamic named range in .NET after row insertion with Aspose.Cells | Set named range to expand with COUNTA after inserting rows using Aspose.Cells | Programmatically adjust Excel named range after inserting rows in C#
// Tags: dynamic named range OFFSET formula Aspose.Cells | insert rows adjust named range C# | Aspose.Cells workbook named range update | Excel dynamic range automation .NET | COUNTA based named range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, inserts three rows at a specified position on Sheet1, redefines the named range "MyDynamicRange" with an OFFSET formula that expands based on the count of non‑empty cells in column A, and saves the updated file as output.xlsx.
class UpdateDynamicNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the target worksheet (by name or index)
            Worksheet sheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name
            if (sheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            // Insert new rows (e.g., insert 3 rows starting at row 5)
            int startRow = 5;          // zero‑based index for Aspose.Cells
            int rowsToInsert = 3;
            sheet.Cells.InsertRows(startRow, rowsToInsert);

            // Retrieve the named range that should stay dynamic
            // Assume the named range is called "MyDynamicRange"
            Name dynamicRange = workbook.Worksheets.Names["MyDynamicRange"];

            if (dynamicRange != null)
            {
                // Re‑define the named range using a dynamic OFFSET formula.
                // This formula always expands to the number of non‑empty cells in column A,
                // starting from A2. Adjust the column/starting cell as needed.
                dynamicRange.RefersTo = "=OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,1)";
            }
            else
            {
                Console.WriteLine("Named range 'MyDynamicRange' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
