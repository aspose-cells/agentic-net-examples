// Title: C# – Apply a Custom Average Formula to a Table’s Totals Row Using Aspose.Cells
// Description: Shows how to build a new workbook, populate rows with item data, turn the range A1:C6 into a ListObject, enable the totals row, mark the Price column for custom calculation, and assign the structured reference =AVERAGE([Price]) through SetCustomTotalsRowFormula before exporting to an XLSX file.
// Keywords: Aspose.Cells C# | SetCustomTotalsRowFormula | Excel table totals row | custom calculation | structured reference formula | average of column | ListObject API | C# spreadsheet automation | totals row custom formula | Aspose.Cells examples
// Common Searches: Aspose.Cells set custom formula in totals row C# | How to calculate average in table totals row with Aspose.Cells | ListObject SetCustomTotalsRowFormula example | C# code for average column in Excel table totals row | structured reference AVERAGE formula Aspose.Cells
// Developer Intent: Generate an Excel file where a table’s totals row displays the mean value of a chosen column via a user‑defined formula.
// Use Cases: Sales ledger that highlights average unit price at the bottom of the price column. | Inventory sheet that reports the mean quantity across items in the totals row. | Budget overview where the average expense per category is shown in the table footer.
// AI Prompts: Create C# Aspose.Cells code that adds a table with a totals row and sets a custom AVERAGE formula for the "Price" column. | Explain the parameters of SetCustomTotalsRowFormula, including A1 vs R1C1 notation and locale settings. | Show how to reference a column by its index when assigning a custom totals‑row formula in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Shows how to build a new workbook, populate rows with item data, turn the range A1:C6 into a ListObject, enable the totals row, mark the Price column for custom calculation, and assign the structured reference =AVERAGE([Price]) through SetCustomTotalsRowFormula before exporting to an XLSX file.
    public class SetCustomTotalsRowAverage
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (header + 5 rows)
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Quantity");
                cells["C1"].PutValue("Price");

                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["C2"].PutValue(1.5);

                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(20);
                cells["C3"].PutValue(0.8);

                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(15);
                cells["C4"].PutValue(2.0);

                cells["A5"].PutValue("Date");
                cells["B5"].PutValue(5);
                cells["C5"].PutValue(3.5);

                cells["A6"].PutValue("Elderberry");
                cells["B6"].PutValue(8);
                cells["C6"].PutValue(4.2);

                // Add a table that includes the data range (A1:C6) and enable totals row
                int tableIndex = sheet.ListObjects.Add("A1", "C6", true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.ShowTotals = true;

                // Choose the column for which we want a custom average in the totals row (e.g., "Price")
                ListColumn priceColumn = table.ListColumns["Price"]; // can also use index 2

                // Set the totals calculation type to Custom
                priceColumn.TotalsCalculation = TotalsCalculation.Custom;

                // Define a custom formula for the totals row using structured reference
                // isR1C1 = false (A1 style), isLocal = false (invariant locale)
                priceColumn.SetCustomTotalsRowFormula("=AVERAGE([Price])", false, false);

                // Ensure output directory exists
                string outputPath = "CustomTotalsRowAverage.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetCustomTotalsRowAverage.Run();
        }
    }
}
