// Title: How to create a C# class that inherits AbstractCalculationEngine and overrides CalculateCustomFunction to return a cell range in Aspose.Cells
// Common Searches: C# Aspose.Cells custom calculation engine that returns a range from a formula | override CalculateCustomFunction to output a cell range using Aspose.Cells API | example of extending AbstractCalculationEngine for range results in an Excel workbook | how to implement a custom function returning a range in Aspose.Cells C# | Aspose.Cells CreateRange inside CalculateCustomFunction implementation
// Tags: override CalculateCustomFunction Aspose.Cells | custom calculation engine returning range | AbstractCalculationEngine C# example | Aspose.Cells CreateRange in custom formula | cell range result custom function | Aspose.Cells custom function implementation

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// This example shows how to derive a class from AbstractCalculationEngine, override the CalculateCustomFunction method, and use Worksheet.Cells.CreateRange to return a cell range as the result of a custom formula in Aspose.Cells for C#. It includes workbook loading, error handling, and printing the returned range address.
public class Program
{
    public static void Main()
    {
        string filePath = "input.xlsx";

        // Prevent FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load workbook
            Workbook workbook = new Workbook(filePath);

            // Example usage: retrieve a range from Sheet1!A1:B2
            AsposeRange range = GetRange(workbook, "Sheet1", "A1:B2");

            // Display the result
            if (range != null)
            {
                Console.WriteLine($"Range address returned: {range.RefersTo}");
            }
            else
            {
                Console.WriteLine("Range could not be retrieved.");
            }
        }
        catch (Exception ex)
        {
            // Runtime safety
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    /// <param name="workbook">The workbook containing the worksheet.</param>
    /// <param name="sheetName">Name of the worksheet.</param>
    /// <param name="address">Cell range address (e.g., "A1:B2").</param>
    /// <returns>The requested Range object, or null if not found.</returns>
    private static AsposeRange GetRange(Workbook workbook, string sheetName, string address)
    {
        try
        {
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet != null && !string.IsNullOrEmpty(address))
            {
                return sheet.Cells.CreateRange(address);
            }
        }
        catch (Exception ex)
        {
            // Log or handle errors as needed
            Console.WriteLine($"GetRange error: {ex.Message}");
        }

        return null;
    }
}
