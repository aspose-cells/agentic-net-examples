using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate some sample data in the range A1:C3
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // -------------------------------------------------
            // Define a workbook‑scoped named range "GlobalRange"
            // -------------------------------------------------
            // Add a new name to the workbook's NameCollection
            int nameIndex = workbook.Worksheets.Names.Add("GlobalRange");
            Name globalName = workbook.Worksheets.Names[nameIndex];

            // Set the reference to the desired range on the worksheet.
            // The leading "=" is required for the RefersTo property.
            // SheetIndex = 0 (default) makes the name workbook‑scoped.
            globalName.RefersTo = "=DataSheet!$A$1:$C$3";

            // Optional: verify the name is visible and has no comment
            globalName.IsVisible = true;
            globalName.Comment = "Workbook‑scoped range covering A1:C3";

            // -------------------------------------------------
            // Retrieve the named range using GetRangeByName
            // -------------------------------------------------
            AsposeRange retrievedRange = workbook.Worksheets.GetRangeByName("GlobalRange");
            if (retrievedRange != null)
            {
                Console.WriteLine($"Retrieved range address: {retrievedRange.Address}");
                // Example: write the sum of the range into cell D1
                sheet.Cells["D1"].Formula = "=SUM(GlobalRange)";
                workbook.CalculateFormula();
                Console.WriteLine($"Sum of GlobalRange: {sheet.Cells["D1"].Value}");
            }
            else
            {
                Console.WriteLine("Named range not found.");
            }

            // Save the workbook as an XLSX file
            workbook.Save("WorkbookScopedNamedRange.xlsx");
        }
    }
}