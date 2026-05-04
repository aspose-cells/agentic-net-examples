using System;
using System.Drawing;
using Aspose.Cells;

namespace AdvancedXlsxDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Units Sold");
            cells["C1"].PutValue("Unit Price");
            cells["D1"].PutValue("Revenue Category");
            cells["E1"].PutValue("Status");

            // Sample rows
            string[] products = { "Laptop", "Phone", "Tablet", "Monitor", "Keyboard", "Mouse", "Printer", "Speaker", "Camera", "Headset" };
            Random rnd = new Random();
            for (int i = 0; i < products.Length; i++)
            {
                int row = i + 2;
                cells[$"A{row}"].PutValue(products[i]);
                cells[$"B{row}"].PutValue(rnd.Next(50, 200));
                cells[$"C{row}"].PutValue(Math.Round(rnd.NextDouble() * 900 + 100, 2));
            }

            // Add formulas
            for (int i = 0; i < products.Length; i++)
            {
                int row = i + 2;
                cells[$"F{row}"].Formula = $"=B{row}*C{row}";
                cells[$"D{row}"].Formula = $"=IF(F{row}>50000,\"High\",IF(F{row}>20000,\"Medium\",\"Low\"))";
            }

            // Header style
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            headerStyle.Borders[BorderType.BottomBorder].Color = Color.Black;

            StyleFlag flag = new StyleFlag { All = true };
            cells.CreateRange("A1:E1").ApplyStyle(headerStyle, flag);

            // Data validation for Status column (E)
            CellArea validationArea = new CellArea
            {
                StartRow = 1,          // Row 2 (zero‑based)
                StartColumn = 4,       // Column E (zero‑based)
                EndRow = products.Length, // Last data row (zero‑based)
                EndColumn = 4
            };

            int validationIndex = sheet.Validations.Add(validationArea);
            Validation statusValidation = sheet.Validations[validationIndex];
            statusValidation.Type = ValidationType.List;
            statusValidation.Operator = OperatorType.Equal;
            statusValidation.Formula1 = "\"Pending,Confirmed,Cancelled\"";
            statusValidation.ShowError = true;
            statusValidation.ErrorTitle = "Invalid Status";
            statusValidation.ErrorMessage = "Please select a value from the list.";

            // Calculate formulas
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "AdvancedDemo.xlsx";
            workbook.Save(outputPath);
            workbook.Dispose();

            Console.WriteLine($"Workbook created and saved to '{outputPath}'.");
        }
    }
}