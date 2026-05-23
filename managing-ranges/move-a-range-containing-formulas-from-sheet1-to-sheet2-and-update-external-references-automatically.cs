using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class MoveRangeWithFormulas
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate Sheet1 with values and formulas
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["B1"].Formula = "=A1+5";          // Internal reference
            sheet1.Cells["B2"].Formula = "=A2*2";          // Internal reference
            sheet1.Cells["C1"].Formula = "=Sheet2!D1+100"; // External reference to Sheet2

            // Define the source range that contains the data and formulas
            AsposeRange sourceRange = sheet1.Cells.CreateRange("A1:C2");

            // Define the destination range on Sheet2 where the range will be moved
            AsposeRange destinationRange = sheet2.Cells.CreateRange("A1:C2");

            // Copy the range (including formulas) to the destination sheet.
            // Formulas with relative references are automatically adjusted.
            destinationRange.CopyData(sourceRange);

            // Optionally clear the original range to complete the "move" operation
            sourceRange.Clear();

            // Recalculate all formulas so that updated references are evaluated
            workbook.CalculateFormula();

            // Save the workbook (ensure the directory exists)
            string outputPath = "MovedRange.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}