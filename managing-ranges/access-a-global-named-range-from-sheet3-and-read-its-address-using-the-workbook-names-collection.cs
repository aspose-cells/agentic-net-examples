using System;
using Aspose.Cells;

namespace Example
{
    class AccessGlobalNamedRange
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet named "Sheet3"
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

                // Populate some cells in Sheet3
                sheet3.Cells["A1"].PutValue("Item1");
                sheet3.Cells["B2"].PutValue("Item2");

                // Define a global named range that refers to Sheet3!A1:B2
                int nameIdx = workbook.Worksheets.Names.Add("GlobalRange");
                Name globalName = workbook.Worksheets.Names[nameIdx];
                globalName.SheetIndex = -1; // -1 indicates global scope
                globalName.RefersTo = "=Sheet3!$A$1:$B$2";

                // Retrieve the global named range
                Name retrievedName = workbook.Worksheets.Names["GlobalRange"];
                Aspose.Cells.Range range = retrievedName.GetRange();
                string address = range.Address;

                Console.WriteLine("Global named range address: " + address);

                // Save the workbook
                string outputPath = "GlobalNamedRangeDemo.xlsx";
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