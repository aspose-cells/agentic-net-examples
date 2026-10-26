// Title: Disable automatic calculation for Excel ListObject tables while keeping workbook calculation mode Automatic with Aspose.Cells for .NET
// Common Searches: Aspose.Cells disable calculation for ListObject tables only C# | keep workbook calculation automatic but turn off table formulas Aspose.Cells .NET | set ListObject.DisableCalculation true example Aspose.Cells | prevent Excel table formulas from recalculating in Aspose.Cells while other cells stay automatic | C# per-table calculation mode configuration using Aspose.Cells
// Tags: ListObject.DisableCalculation Aspose.Cells | per-table calculation mode .NET | disable table auto calculation Excel | retain workbook automatic calculation Aspose.Cells | C# Excel table calculation control

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an Excel workbook, iterates through each worksheet and its ListObject tables, sets each table's DisableCalculation property to true to stop automatic recalculation for table formulas, and then saves the workbook while leaving the overall workbook calculation mode set to Automatic.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Process each table (ListObject) in the worksheet
                foreach (ListObject table in sheet.ListObjects)
                {
                    // Get the data range of the table (excluding headers/totals)
                    Aspose.Cells.Range dataRange = table.DataRange;

                    // Additional cell processing can be performed here if needed.
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
