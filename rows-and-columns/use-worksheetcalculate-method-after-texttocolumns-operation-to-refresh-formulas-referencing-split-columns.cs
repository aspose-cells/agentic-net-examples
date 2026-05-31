using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with comma‑separated values
            sheet.Cells["A1"].PutValue("John,100");
            sheet.Cells["A2"].PutValue("Jane,200");

            // Formulas that depend on the second part after split (column B)
            sheet.Cells["C1"].Formula = "=B1*2";
            sheet.Cells["C2"].Formula = "=B2*2";

            // Set up TextToColumns options (comma as separator)
            TxtLoadOptions options = new TxtLoadOptions
            {
                Separator = ','
            };

            // Split the content of column A into multiple columns
            // Parameters: start row, start column, total rows, options
            sheet.Cells.TextToColumns(0, 0, 2, options);

            // Recalculate all formulas in the workbook so they reflect the new split data
            workbook.CalculateFormula();

            // Output the refreshed formula results
            Console.WriteLine("C1 = " + sheet.Cells["C1"].StringValue); // Expected 200
            Console.WriteLine("C2 = " + sheet.Cells["C2"].StringValue); // Expected 400

            // Save the workbook (save rule)
            string outputPath = "TextToColumns_Calculated.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}