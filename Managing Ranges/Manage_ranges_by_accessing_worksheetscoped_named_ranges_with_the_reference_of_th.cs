using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace WorksheetScopedNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add two worksheets: one for data, one for summary
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "DataSheet";

            Worksheet summarySheet = workbook.Worksheets.Add("Summary");

            // Populate some sample data in DataSheet (A1:B3)
            dataSheet.Cells["A1"].PutValue("Item");
            dataSheet.Cells["B1"].PutValue("Quantity");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("Orange");
            dataSheet.Cells["B3"].PutValue(15);

            // Create a named range that is scoped to DataSheet
            int nameIndex = workbook.Worksheets.Names.Add("MyData");
            Name scopedName = workbook.Worksheets.Names[nameIndex];

            // Set the scope to the DataSheet (zero‑based index)
            scopedName.SheetIndex = dataSheet.Index;

            // Define the range reference (absolute address)
            scopedName.RefersTo = $"={dataSheet.Name}!$A$1:$B$3";

            // ------------------------------
            // Access the worksheet‑scoped named range
            // ------------------------------

            // Method 1: Retrieve via the Name object and GetRange()
            AsposeRange rangeViaName = scopedName.GetRange();
            Console.WriteLine($"Method 1 - Range address: {rangeViaName.Address}");
            Console.WriteLine($"Method 1 - Belongs to worksheet: {rangeViaName.Worksheet.Name}");

            // Method 2: Retrieve via WorksheetCollection.GetRangeByName with sheet index
            AsposeRange rangeViaCollection = workbook.Worksheets.GetRangeByName("MyData", dataSheet.Index, false);
            if (rangeViaCollection != null)
            {
                Console.WriteLine($"Method 2 - Range address: {rangeViaCollection.Address}");
                Console.WriteLine($"Method 2 - Belongs to worksheet: {rangeViaCollection.Worksheet.Name}");
            }
            else
            {
                Console.WriteLine("Method 2 - Named range not found for the specified sheet.");
            }

            // Use the named range in a formula on the Summary sheet
            summarySheet.Cells["A1"].Formula = "=SUM(MyData)";
            workbook.CalculateFormula();

            Console.WriteLine($"Sum of MyData (displayed on Summary!A1): {summarySheet.Cells["A1"].Value}");

            // Save the workbook in XLSX format
            workbook.Save("WorksheetScopedNamedRange.xlsx");
        }
    }
}