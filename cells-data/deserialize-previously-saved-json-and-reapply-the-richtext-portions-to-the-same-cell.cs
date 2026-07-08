using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsRichTextJsonDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and add rich‑text formatting to a cell
            // ------------------------------------------------------------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cell sourceCell = sourceSheet.Cells["A1"];
            sourceCell.PutValue("Hello World");

            // Apply bold to "Hello"
            FontSetting boldPart = sourceCell.Characters(0, 5);
            boldPart.Font.IsBold = true;

            // Apply red color to "World"
            FontSetting redPart = sourceCell.Characters(6, 5);
            redPart.Font.Color = Color.Red;

            // Serialize the cell (including rich‑text info) to JSON
            string cellJson = sourceCell.ToJson();

            // ------------------------------------------------------------
            // 2. Create a new workbook and import the previously saved JSON
            // ------------------------------------------------------------
            Workbook targetWorkbook = new Workbook();
            Worksheet targetSheet = targetWorkbook.Worksheets[0];

            // JsonLayoutOptions can be default; no special settings required
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions();

            // Import the JSON into the target worksheet starting at cell A1
            JsonUtility.ImportData(cellJson, targetSheet.Cells, 0, 0, layoutOptions);

            // ------------------------------------------------------------
            // 3. Save both workbooks for verification
            // ------------------------------------------------------------
            sourceWorkbook.Save("SourceRichText.xlsx", SaveFormat.Xlsx);
            targetWorkbook.Save("TargetFromJson.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Rich‑text cell serialized to JSON and re‑imported successfully.");
        }
    }
}