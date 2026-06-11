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
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some initial data in column O (index 14)
                sheet.Cells["O1"].PutValue("Header");
                sheet.Cells["O2"].PutValue(10);
                sheet.Cells["O3"].PutValue(20);
                sheet.Cells["O4"].PutValue(30);

                // ---------- Create a dynamic named range ----------
                // The formula uses OFFSET together with COUNTA to expand automatically
                // as new non‑empty rows are added to column O.
                int nameIndex = workbook.Worksheets.Names.Add("DynamicO");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                // Example formula: =OFFSET(Sheet1!$O$2,0,0,COUNTA(Sheet1!$O:$O)-1,1)
                dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$O$2,0,0,COUNTA({sheet.Name}!$O:$O)-1,1)";

                // ---------- Verify the named range before adding new rows ----------
                workbook.CalculateFormula(); // Ensure the formula is evaluated
                AsposeRange rangeBefore = dynamicName.GetRange();
                Console.WriteLine($"Range before adding rows: {rangeBefore.Address} (Rows: {rangeBefore.RowCount})");

                // ---------- Add new rows to column O ----------
                sheet.Cells["O5"].PutValue(40);
                sheet.Cells["O6"].PutValue(50);

                // Re‑calculate to let the dynamic formula update
                workbook.CalculateFormula();

                // Retrieve the updated range
                AsposeRange rangeAfter = dynamicName.GetRange();
                Console.WriteLine($"Range after adding rows: {rangeAfter.Address} (Rows: {rangeAfter.RowCount})");

                // ---------- Save the workbook ----------
                string outputPath = "DynamicNamedRange.xlsx";
                // Ensure we can write to the target location
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
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