// Title: Convert an Aspose.Cells ListObject (Excel table) back to a regular range while preserving formulas in C#
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, locates a ListObject, converts it back to a normal range, and verifies that formulas stay intact. | Provide a complete C# example that creates an Excel table, sets a formula in a column, transforms the table into a regular cell range, and saves the workbook.
// Common Searches: Aspose.Cells C# convert Excel table to range without losing formulas | ListObject ConvertToRange keep formulas Aspose.Cells example | how to remove a table in Aspose.Cells while retaining cell formulas | C# Aspose.Cells preserve formulas when converting ListObject to range
// Tags: Aspose.Cells ListObject to range conversion | preserve cell formulas during Excel table removal | C# delete Excel table while keeping formulas | convert Excel ListObject back to regular cells | save workbook after converting table to range C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads or creates a workbook, optionally adds a ListObject with a formula, uses the ListObject conversion method to turn the table back into a normal cell range while keeping formulas unchanged, and then saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if present; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            Worksheet sheet = workbook.Worksheets[0];

            // If a ListObject (table) exists, convert it back to a range.
            if (sheet.ListObjects.Count > 0)
            {
                ListObject listObject = sheet.ListObjects[0];
                listObject.ConvertToRange(); // preserves formulas
            }
            else
            {
                // Create a sample table to demonstrate conversion when none exists.
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["C1"].PutValue("Header3");
                sheet.Cells["A2"].PutValue(1);
                // Set formula for B2 cell.
                sheet.Cells["B2"].Formula = "=A2*2";
                sheet.Cells["C2"].PutValue(3);

                // Add a ListObject (table) over the range A1:C2.
                int tableIndex = sheet.ListObjects.Add(0, 0, 2, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Convert the newly created table back to a regular range.
                table.ConvertToRange();
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
