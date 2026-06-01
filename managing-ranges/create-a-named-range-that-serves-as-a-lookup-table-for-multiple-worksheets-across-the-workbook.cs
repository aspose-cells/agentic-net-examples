using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeLookup
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet that will contain the lookup table
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate the lookup table (Item | Price)
                dataSheet.Cells["A1"].PutValue("Item");
                dataSheet.Cells["B1"].PutValue("Price");
                dataSheet.Cells["A2"].PutValue("Apple");
                dataSheet.Cells["B2"].PutValue(1.20);
                dataSheet.Cells["A3"].PutValue("Banana");
                dataSheet.Cells["B3"].PutValue(0.80);
                dataSheet.Cells["A4"].PutValue("Cherry");
                dataSheet.Cells["B4"].PutValue(2.50);
                dataSheet.Cells["A5"].PutValue("Date");
                dataSheet.Cells["B5"].PutValue(3.00);

                // Create a global named range that refers to the lookup table
                int nameIndex = workbook.Worksheets.Names.Add("ItemPriceTable");
                Name lookupName = workbook.Worksheets.Names[nameIndex];
                // Global scope: SheetIndex = 0 (default)
                lookupName.RefersTo = "=Data!$A$2:$B$5";

                // Add another worksheet where the lookup will be used
                Worksheet reportSheet = workbook.Worksheets.Add("Report");

                // Example: retrieve price of "Cherry" using VLOOKUP and the named range
                reportSheet.Cells["A1"].PutValue("Lookup Item");
                reportSheet.Cells["B1"].PutValue("Price");
                reportSheet.Cells["A2"].PutValue("Cherry");
                // VLOOKUP formula referencing the global named range
                reportSheet.Cells["B2"].Formula = "=VLOOKUP(A2, ItemPriceTable, 2, FALSE)";

                // Optional: demonstrate retrieving the named range programmatically
                // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
                Aspose.Cells.Range retrievedRange = workbook.Worksheets.GetRangeByName("ItemPriceTable");
                if (retrievedRange != null)
                {
                    Console.WriteLine($"Named range '{retrievedRange.Name}' refers to: {retrievedRange.RefersTo}");
                }

                // Calculate formulas so that the VLOOKUP result is evaluated
                workbook.CalculateFormula();

                // Output the lookup result to console
                Console.WriteLine($"Price of {reportSheet.Cells["A2"].StringValue}: {reportSheet.Cells["B2"].Value}");

                // Save the workbook
                string outputPath = "NamedRangeLookupDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}