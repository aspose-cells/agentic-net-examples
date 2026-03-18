using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string inputPath = @"C:\Data\SampleInput.xlsx";
            string outputPath = @"C:\Data\ModifiedOutput.xlsx";

            // Ensure the output directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Access the first worksheet and modify cells
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["B2"].PutValue("Updated Value");
            sheet.Cells["C5"].PutValue(12345);

            // Insert a new worksheet and add data
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "Summary";

            newSheet.Cells["A1"].PutValue("Item");
            newSheet.Cells["B1"].PutValue("Quantity");
            newSheet.Cells["C1"].PutValue("Price");

            newSheet.Cells["A2"].PutValue("Apples");
            newSheet.Cells["B2"].PutValue(50);
            newSheet.Cells["C2"].PutValue(0.75);

            newSheet.Cells["A3"].PutValue("Bananas");
            newSheet.Cells["B3"].PutValue(30);
            newSheet.Cells["C3"].PutValue(0.55);

            // Apply style to header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            newSheet.Cells.CreateRange("A1", "C1")
                .ApplyStyle(headerStyle, new StyleFlag { Font = true, CellShading = true });

            // Save the modified workbook
            workbook.Save(outputPath);
            workbook.Dispose();

            Console.WriteLine($"Workbook has been modified and saved to: {outputPath}");
        }
    }
}