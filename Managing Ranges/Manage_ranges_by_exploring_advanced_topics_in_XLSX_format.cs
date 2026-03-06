using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsAdvancedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Quantity");
            cells["D1"].PutValue("Price");

            string[,] data = {
                { "Fruit", "Apple", "10", "0.5" },
                { "Fruit", "Banana", "20", "0.3" },
                { "Vegetable", "Carrot", "15", "0.2" },
                { "Vegetable", "Tomato", "12", "0.4" }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    cells[i + 2, j].PutValue(data[i, j]); // rows start at index 2 (A3)
                }
            }

            // Create a named range covering the data block (A2:D5)
            Aspose.Cells.Range dataRange = cells.CreateRange("A2", "D5");
            dataRange.Name = "SalesData";

            // Use CurrentRegion property
            Aspose.Cells.Range singleCell = cells.CreateRange("B3", "B3");
            Aspose.Cells.Range region = singleCell.CurrentRegion; // should be A2:D5
            Console.WriteLine($"CurrentRegion address: {region.Address}");

            // Copy the range to a new location with PasteOptions
            Aspose.Cells.Range destination = cells.CreateRange("F2", "I5"); // copy to columns F-I
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All
            };
            destination.Copy(dataRange, pasteOptions);

            // Add the range to the runtime Ranges collection
            cells.Ranges.Add(dataRange);

            // Apply style to the entire region
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightYellow;
            style.Pattern = BackgroundType.Solid;
            region.SetStyle(style);

            // Export the named range to JSON
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = false
            };
            string json = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);
            File.WriteAllText("SalesData.json", json);
            Console.WriteLine("Exported range to JSON file.");

            // Save workbook as XLSX
            workbook.Save("AdvancedRangeDemo.xlsx");

            // Save only a specific area using XmlSaveOptions
            XmlSaveOptions xmlOptions = new XmlSaveOptions
            {
                ExportArea = new CellArea { StartRow = 1, EndRow = 5, StartColumn = 0, EndColumn = 3 } // A2:D6
            };
            workbook.Save("PartialExport.xml", xmlOptions);
            Console.WriteLine("Saved partial area to XML.");

            // Save as SpreadsheetML with XLS limits
            SpreadsheetML2003SaveOptions xlLimits = new SpreadsheetML2003SaveOptions
            {
                LimitAsXls = true
            };
            workbook.Save("LimitedSpreadsheet.xml", xlLimits);
            Console.WriteLine("Saved SpreadsheetML with XLS limits.");

            // Cleanup
            workbook.Dispose();
        }
    }
}