// Title: Replace #DIV/0! error values with zero in an Aspose.Cells workbook using C#
// Common Searches: C# Aspose.Cells replace #DIV/0! error with 0 in Excel file | how to set error cells to zero after formula calculation using Aspose.Cells .NET | Aspose.Cells C# change division by zero error value to numeric | remove #DIV/0! from generated workbook with Aspose.Cells API | replace Excel error values programmatically using Aspose.Cells C#
// Tags: replace Excel error values with zero Aspose.Cells C# | handle division by zero error Aspose.Cells workbook | set cell error type to numeric Aspose.Cells | calculate formulas and substitute #DIV/0! Aspose.Cells | C# Aspose.Cells error value substitution

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds a formula that causes a #DIV/0! error, recalculates the sheet, detects the error cell, replaces the error value with zero, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set up data that will cause a division by zero error
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(0);   // divisor = 0
            sheet.Cells["A2"].PutValue(10);  // dividend = 10
            sheet.Cells["B1"].Formula = "A2/A1"; // division by zero

            // Recalculate formulas so the error value is generated
            workbook.CalculateFormula();

            // Replace error values (e.g., #DIV/0!) with zero
            Cell resultCell = sheet.Cells["B1"];
            if (resultCell.Type == CellValueType.IsError)
            {
                resultCell.PutValue(0);
            }

            // Define output file path
            string outputPath = "output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
