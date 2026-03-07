using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the default worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            // Add a total formula for quantity and price
            sheet.Cells["B4"].Formula = "=SUM(B2:B3)";
            sheet.Cells["C4"].Formula = "=SUMPRODUCT(B2:B3,C2:C3)";

            // Apply a simple style to the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the first row
            sheet.Cells.CreateRange("A1:C1").ApplyStyle(headerStyle, new StyleFlag
            {
                FontBold = true,
                CellShading = true
            });

            // Calculate formulas so that results are stored
            workbook.CalculateFormula();

            // Save the workbook to disk (lifecycle: save)
            string outputPath = "SampleReport.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}