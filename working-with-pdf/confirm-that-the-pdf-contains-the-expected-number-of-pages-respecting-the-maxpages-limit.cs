// Title: Validate that an Excel workbook contains no more than a specified number of worksheets using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells check worksheet count against a maximum limit | how to enforce a maximum number of sheets in an Excel file with Aspose.Cells | validate Excel workbook sheet count programmatically using .NET | Aspose.Cells example to limit number of worksheets in a workbook | C# code to verify Excel file does not exceed allowed sheet count
// Tags: worksheet count validation Aspose.Cells | limit Excel workbook sheets C# | Aspose.Cells workbook sheet enumeration | Excel file sheet count constraint .NET | validate workbook sheet limit programmatically

using Aspose.Cells;
using System;
using System.IO;

// The example shows how to use Aspose.Cells for .NET to load an Excel workbook, retrieve its Worksheets.Count, and compare it with a user‑defined maximum. The ValidateSheetCount method returns true when the sheet count is within the limit, and the sample Main method demonstrates handling missing files, reporting the result, and catching exceptions.
public class ExcelSheetValidator
{
    // Checks if the Excel file at the given path has a worksheet count within the specified limit.
    public static bool ValidateSheetCount(string excelFilePath, int maxSheets)
    {
        try
        {
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"File not found: {excelFilePath}");
                return false;
            }

            // Load the Excel workbook.
            Workbook workbook = new Workbook(excelFilePath);
            // Retrieve the total number of worksheets.
            int sheetCount = workbook.Worksheets.Count;
            // Return true if the sheet count does not exceed the limit.
            return sheetCount <= maxSheets;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error validating sheet count: {ex.Message}");
            return false;
        }
    }

    // Example usage.
    public static void Main()
    {
        string excelPath = "sample.xlsx";
        int maxSheets = 5;

        bool isWithinLimit = ValidateSheetCount(excelPath, maxSheets);

        try
        {
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                return;
            }

            Workbook workbook = new Workbook(excelPath);
            int totalSheets = workbook.Worksheets.Count;

            if (isWithinLimit)
            {
                Console.WriteLine($"Excel file has {totalSheets} worksheets, which is within the allowed maximum of {maxSheets}.");
            }
            else
            {
                Console.WriteLine($"Excel file has {totalSheets} worksheets, exceeding the allowed maximum of {maxSheets}.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing Excel file: {ex.Message}");
        }
    }
}
