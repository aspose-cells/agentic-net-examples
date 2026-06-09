using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicNamedRange
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate initial data in the first row (A1, B1, C1)
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["C1"].PutValue("Header3");

                // Create a dynamic named range that expands horizontally.
                // The formula uses OFFSET with COUNTA to count non‑empty cells in row 1.
                // It always starts at A1, has 1 row height, and width equals the number of filled cells.
                int nameIndex = workbook.Worksheets.Names.Add("MyDynamicRange");
                Name dynName = workbook.Worksheets.Names[nameIndex];
                dynName.RefersTo = "=OFFSET(DataSheet!$A$1,0,0,1,COUNTA(DataSheet!$1:$1))";

                // Verify the range before adding new columns
                AsposeRange initialRange = dynName.GetRange();
                Console.WriteLine($"Initial range address: {initialRange.Address} (Columns: {initialRange.ColumnCount})");

                // Add two more columns to demonstrate automatic expansion
                sheet.Cells["D1"].PutValue("Header4");
                sheet.Cells["E1"].PutValue("Header5");

                // Recalculate formulas so that COUNTA updates
                workbook.CalculateFormula();

                // Retrieve the named range again; it should now include the new columns
                AsposeRange expandedRange = dynName.GetRange();
                Console.WriteLine($"Expanded range address: {expandedRange.Address} (Columns: {expandedRange.ColumnCount})");

                // Use the dynamic named range in a formula (e.g., concatenate headers)
                sheet.Cells["A2"].Formula = "=TEXTJOIN(\",\",TRUE,MyDynamicRange)";
                workbook.CalculateFormula();
                Console.WriteLine($"Concatenated headers: {sheet.Cells["A2"].StringValue}");

                // Define output file path
                string outputPath = "DynamicNamedRangeDemo.xlsx";

                // Ensure we can write to the target location
                try
                {
                    if (File.Exists(outputPath))
                    {
                        File.Delete(outputPath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not delete existing file. {ex.Message}");
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}