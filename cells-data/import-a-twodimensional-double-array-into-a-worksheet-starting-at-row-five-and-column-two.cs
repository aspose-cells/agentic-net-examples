// Title: Import a 2‑D double array into a worksheet at row 5, column 2 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# ImportTwoDimensionArray starting at row 5 column 2 | How to load a double[,] into an Excel sheet at a specific cell range with Aspose.Cells | C# example for inserting a 2D array into a worksheet and overwriting existing cells | Using Worksheet.Cells.ImportTwoDimensionArray to write numeric data in .NET | Saving an Excel file after importing a 2D array with Aspose.Cells
// Tags: Worksheet.Cells.ImportTwoDimensionArray method | import 2d double array Aspose.Cells C# | specify start row and column for data import | overwrite existing cells during array import | save workbook after importing data Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook, accesses the first Worksheet, defines a 3×3 double array (as object[,]), imports it beginning at row 5, column 2 (zero‑based indices 4 and 1) with the overwrite flag set to false, and then saves the workbook as 'ImportedArray.xlsx', handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Two‑dimensional data to import (as object[,] because Aspose.Cells expects object array)
            object[,] data = new object[,]
            {
                { 1.1, 2.2, 3.3 },
                { 4.4, 5.5, 6.6 },
                { 7.7, 8.8, 9.9 }
            };

            // Import the array starting at row 5 (index 4) and column 2 (index 1)
            // The last parameter 'false' indicates that existing cells will be overwritten
            worksheet.Cells.ImportTwoDimensionArray(data, 4, 1, false);

            // Save the workbook (lifecycle rule)
            workbook.Save("ImportedArray.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
