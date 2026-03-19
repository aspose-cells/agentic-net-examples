using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add some data
            // -------------------------------------------------
            // Use the default constructor (Workbook())
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set worksheet name
            sheet.Name = "DataSheet";

            // Populate some cells with sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            // Save the newly created workbook to disk
            // Use the Save(string) overload
            string createdPath = "CreatedWorkbook.xlsx";
            workbook.Save(createdPath);
            Console.WriteLine($"Workbook created and saved to '{createdPath}'.");

            // -------------------------------------------------
            // 2. Load the workbook from file and modify it
            // -------------------------------------------------
            // Use the constructor that accepts a file path (Workbook(string))
            Workbook loadedWorkbook = new Workbook(createdPath);

            // Access the same worksheet (index 0)
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Insert a new column for "Total" after the Price column
            loadedSheet.Cells.InsertColumn(3); // Column index is zero‑based, so 3 = column D

            // Add header for the new column
            loadedSheet.Cells["D1"].PutValue("Total");

            // Calculate total = Quantity * Price for each row
            for (int row = 2; row <= 3; row++) // rows 2 and 3 contain data
            {
                // Build a simple formula string, e.g., =B2*C2
                string quantityCell = $"B{row}";
                string priceCell = $"C{row}";
                string formula = $"={quantityCell}*{priceCell}";
                loadedSheet.Cells[$"D{row}"].Formula = formula;
            }

            // Recalculate formulas so that the Total column has values
            loadedWorkbook.CalculateFormula();

            // Demonstrate the Replace method: replace the product name "Apple" with "Green Apple"
            // This replaces all occurrences of the placeholder text in the worksheet
            loadedWorkbook.Replace("Apple", "Green Apple");

            // Save the modified workbook to a new file
            string modifiedPath = "ModifiedWorkbook.xlsx";
            loadedWorkbook.Save(modifiedPath);
            Console.WriteLine($"Workbook loaded, modified, and saved to '{modifiedPath}'.");

            // -------------------------------------------------
            // 3. Export the modified workbook to PDF using ConversionUtility
            // -------------------------------------------------
            // Use the static Convert method from Aspose.Cells.Utility.ConversionUtility
            // Convert(string source, string saveAs)
            string pdfPath = "ModifiedWorkbook.pdf";
            Aspose.Cells.Utility.ConversionUtility.Convert(modifiedPath, pdfPath);
            Console.WriteLine($"Workbook converted to PDF at '{pdfPath}'.");
        }
    }
}