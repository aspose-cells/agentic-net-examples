using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDemo
{
    class TableStyleCopyDemo
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet (source sheet)
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate source sheet with sample data
            sourceSheet.Cells["A1"].PutValue("Product");
            sourceSheet.Cells["B1"].PutValue("Price");
            sourceSheet.Cells["A2"].PutValue("Apple");
            sourceSheet.Cells["B2"].PutValue(1.2);
            sourceSheet.Cells["A3"].PutValue("Banana");
            sourceSheet.Cells["B3"].PutValue(0.8);

            // Add a table to the source sheet
            int srcTableIndex = sourceSheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject srcTable = sourceSheet.ListObjects[srcTableIndex];
            srcTable.ShowTableStyleFirstColumn = true;
            srcTable.ShowTableStyleLastColumn = true;

            // Apply a built‑in table style to the source table
            // (Custom table styles require the TableStyles collection which may not be available in all versions)
            srcTable.TableStyleName = "TableStyleMedium2";

            // Add a second worksheet (destination sheet) and populate it with similar data
            Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            destSheet.Name = "Destination";

            destSheet.Cells["A1"].PutValue("Product");
            destSheet.Cells["B1"].PutValue("Price");
            destSheet.Cells["A2"].PutValue("Orange");
            destSheet.Cells["B2"].PutValue(1.5);
            destSheet.Cells["A3"].PutValue("Grape");
            destSheet.Cells["B3"].PutValue(2.0);

            // Add a table to the destination sheet
            int destTableIndex = destSheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject destTable = destSheet.ListObjects[destTableIndex];
            destTable.ShowTableStyleFirstColumn = true;
            destTable.ShowTableStyleLastColumn = true;

            // Copy the style from the source table to the destination table
            destTable.TableStyleName = srcTable.TableStyleName;

            // Save the workbook
            string outputPath = "TableStyleCopied.xlsx";
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
    }
}