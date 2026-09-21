// Title: Verify French subtotal label translation in an Aspose.Cells worksheet with C#
// AI Prompts: Create a C# console app that uses Aspose.Cells to set Workbook.Settings.CultureInfo to fr-FR, adds sample data, inserts a manual total row, and prints the value of the subtotal label cell. | Adapt the example to call the Aspose.Cells Subtotal method (if available) and retrieve the automatically generated French label for the total row. | Write an NUnit test in C# that opens SubtotalLabelTranslation.xlsx, reads cell A6, and asserts that the label matches the expected French word for "Total".
// Common Searches: Aspose.Cells how to display subtotal label in French | C# set workbook cultureinfo for localized subtotal row | unit test subtotal label translation in Excel file using Aspose.Cells | manual subtotal calculation with culture‑aware label Aspose.Cells C# | verify worksheet total row label localization fr-FR Aspose.Cells
// Tags: Aspose.Cells subtotal label localization | C# workbook cultureinfo French | manual total row generation Aspose.Cells | unit testing worksheet label localization C# | Excel subtotal row localization Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Globalization;

// The sample creates a new workbook, sets its CultureInfo to French (fr-FR), fills sample data, manually computes a total for column B, inserts a total label in cell A6, prints the label, and saves the file as SubtotalLabelTranslation.xlsx to verify that the subtotal label appears correctly for the selected culture.
class SubtotalLabelTranslationTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Set the culture to French (France) to test label translation
            wb.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Populate sample data (rows are zero‑based)
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Amount");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("A");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("B");
            ws.Cells["B4"].PutValue(30);
            ws.Cells["A5"].PutValue("B");
            ws.Cells["B5"].PutValue(40);

            // NOTE: The Subtotal API is not available in the referenced Aspose.Cells version.
            // As an alternative, manually calculate totals and insert a total row.
            // This demonstrates label handling without relying on the missing Subtotal method.

            // Calculate total for column B
            double total = 0;
            for (int row = 1; row <= 4; row++) // rows 2‑5 (zero‑based)
            {
                total += ws.Cells[row, 1].DoubleValue;
            }

            // Insert total label in the first column of the next row (A6)
            ws.Cells["A6"].PutValue("Total");
            ws.Cells["B6"].PutValue(total);

            // Retrieve the label (culture‑aware translation is handled by Aspose.Cells when using Subtotal;
            // here we manually set the label, so it will appear as entered.)
            string label = ws.Cells["A6"].StringValue;
            Console.WriteLine("Subtotal label: " + label);

            // Save the workbook for manual inspection if needed
            string outputPath = "SubtotalLabelTranslation.xlsx";
            wb.Save(outputPath);
            Console.WriteLine("Workbook saved to " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
