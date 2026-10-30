// Title: How to validate a German subtotal label in an Aspose.Cells worksheet using C#
// Common Searches: how to insert a subtotal row with a German label using Aspose.Cells in C# | C# Aspose.Cells verify that a custom subtotal label appears correctly in Excel | testing localized subtotal label in an Aspose.Cells generated workbook | read cell string value after inserting subtotal row with Aspose.Cells C# | validate German worksheet label translation in saved Excel file using Aspose.Cells
// Tags: Aspose.Cells insert subtotal row C# | Aspose.Cells localized subtotal label | C# verify cell string value Aspose.Cells | German worksheet label validation Aspose.Cells | Excel workbook save custom subtotal label C#

using System;
using Aspose.Cells;

// The example creates a workbook, fills sample data, inserts a subtotal row, sets the label to the German word "Zwischensumme", reads the cell back to confirm the translation, outputs the result, and saves the file as SubtotalLabelTranslationTest.xlsx.
class SubtotalLabelTranslationTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            string[] categories = { "Food", "Food", "Transport", "Transport", "Utilities", "Utilities" };
            double[] amounts = { 120.5, 80.75, 50.0, 30.25, 200.0, 150.0 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]);   // Column A (zero‑based row index)
                cells[i + 1, 1].PutValue(amounts[i]);     // Column B
            }

            // Manually insert a subtotal row after the first group (Food)
            // Row index after the first two data rows (zero‑based)
            int subtotalRowIndex = 3;
            sheet.Cells.InsertRows(subtotalRowIndex, 1);
            sheet.Cells[subtotalRowIndex, 0].PutValue("Zwischensumme");
            double foodSum = amounts[0] + amounts[1];
            sheet.Cells[subtotalRowIndex, 1].PutValue(foodSum);

            // Verify that the subtotal label appears correctly
            string label = sheet.Cells[subtotalRowIndex, 0].StringValue;
            Console.WriteLine("Expected label: Zwischensumme");
            Console.WriteLine("Actual label  : " + label);
            Console.WriteLine(label == "Zwischensumme"
                ? "Label translation succeeded."
                : "Label translation failed.");

            // Save the workbook (optional, for visual inspection)
            string outputPath = "SubtotalLabelTranslationTest.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
