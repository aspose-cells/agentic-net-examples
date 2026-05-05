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

            // Access the default worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Populate sample data
            // -------------------------------------------------
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Quantity");
            cells["D1"].PutValue("Unit Price");
            cells["E1"].PutValue("Total Price");

            string[] products = { "Laptop", "Phone", "Tablet", "Monitor", "Keyboard" };
            string[] categories = { "Electronics", "Electronics", "Electronics", "Electronics", "Accessories" };
            int[] quantities = { 5, 10, 7, 3, 15 };
            double[] unitPrices = { 1200.0, 800.0, 450.0, 300.0, 25.0 };

            for (int i = 0; i < products.Length; i++)
            {
                int row = i + 2;
                cells[$"A{row}"].PutValue(products[i]);
                cells[$"B{row}"].PutValue(categories[i]);
                cells[$"C{row}"].PutValue(quantities[i]);
                cells[$"D{row}"].PutValue(unitPrices[i]);
            }

            // -------------------------------------------------
            // 2. Apply a custom style to the header row
            // -------------------------------------------------
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;

            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 164; // Built‑in currency format

            StyleFlag headerFlag = new StyleFlag { All = true };
            cells.CreateRange("A1:E1").ApplyStyle(headerStyle, headerFlag);

            StyleFlag currencyFlag = new StyleFlag { NumberFormat = true };
            cells.CreateRange("D2:D6").ApplyStyle(currencyStyle, currencyFlag);
            cells.CreateRange("E2:E6").ApplyStyle(currencyStyle, currencyFlag);

            // -------------------------------------------------
            // 3. Add data validation (list) to the Category column (B)
            // -------------------------------------------------
            ValidationCollection validations = sheet.Validations;
            CellArea validationArea = new CellArea
            {
                StartRow = 1,   // B2 (zero‑based)
                StartColumn = 1,
                EndRow = 5,     // B6
                EndColumn = 1
            };
            int vIdx = validations.Add(validationArea);
            Validation validation = validations[vIdx];
            validation.Type = ValidationType.List;
            validation.Formula1 = "\"Electronics,Accessories,Software\"";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Category";
            validation.ErrorMessage = "Please select a category from the predefined list.";
            validation.InCellDropDown = true;

            // -------------------------------------------------
            // 4. Insert a formula for Total Price (Quantity * Unit Price)
            // -------------------------------------------------
            for (int i = 2; i <= 6; i++)
            {
                cells[$"E{i}"].Formula = $"=C{i}*D{i}";
            }

            // -------------------------------------------------
            // 5. Insert formulas that calculate the total quantity per category
            // -------------------------------------------------
            cells["G1"].PutValue("Category");
            cells["H1"].PutValue("Total Quantity");
            cells["G2"].PutValue("Electronics");
            cells["G3"].PutValue("Accessories");
            cells["G4"].PutValue("Software");

            for (int i = 2; i <= 4; i++)
            {
                cells[$"H{i}"].Formula = $"=SUMIFS(C2:C6,B2:B6,G{i})";
            }

            // -------------------------------------------------
            // 6. Calculate all formulas so that values are stored
            // -------------------------------------------------
            workbook.CalculateFormula();

            // -------------------------------------------------
            // 7. Save the workbook
            // -------------------------------------------------
            workbook.Save("AdvancedOperations.xlsx", SaveFormat.Xlsx);
            workbook.Dispose();

            Console.WriteLine("Workbook 'AdvancedOperations.xlsx' created successfully.");
        }
    }
}