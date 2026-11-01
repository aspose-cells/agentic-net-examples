// Title: How to verify that an Excel worksheet has at least one data row and a shape using Aspose.Cells for .NET
// AI Prompts: Write a C# function that receives an Aspose.Cells Worksheet object and returns true only if the worksheet contains at least one populated row (MaxDataRow >= 0) and one or more shapes (Shapes.Count > 0). | Create a C# console program that loads a workbook with Aspose.Cells, calls the above function on a selected worksheet, and prints a message indicating whether both data rows and shapes are present.
// Common Searches: aspnet check worksheet for both data rows and pictures using Aspose.Cells | c# determine if Excel sheet contains any shapes and non‑empty rows with Aspose.Cells | how to use MaxDataRow and Shapes.Count to detect mixed content in an Excel worksheet | sample code to validate worksheet has data and charts in Aspose.Cells for .NET | detect mixed content (rows and drawings) in Excel file programmatically
// Tags: detect worksheet data rows Aspose.Cells | verify worksheet shapes Aspose.Cells | mixed content detection Excel .NET | MaxDataRow shape count validation | worksheet analysis Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Provides a HasDataRowsAndShape method that returns true when the supplied Worksheet has at least one data row (MaxDataRow >= 0) and contains one or more shapes (Shapes.Count > 0). The example program loads a workbook, selects the first worksheet, invokes the method, and outputs the result.
public class WorksheetAnalyzer
{
    /// <param name="sheet">The worksheet to analyze.</param>
    /// <returns>True if both conditions are met; otherwise false.</returns>
    public bool HasDataRowsAndShape(Worksheet sheet)
    {
        // A data row exists if MaxDataRow is non‑negative.
        bool hasDataRow = sheet.Cells.MaxDataRow >= 0;

        // A shape exists if the worksheet contains any shapes (pictures, charts, etc.).
        bool hasShape = sheet.Shapes.Count > 0;

        return hasDataRow && hasShape;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Path to the Excel file to analyze.
            string filePath = "Sample.xlsx";

            // Ensure the file exists before attempting to load it.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(filePath);

            // Use the first worksheet for analysis (adjust as needed).
            Worksheet sheet = workbook.Worksheets[0];

            WorksheetAnalyzer analyzer = new WorksheetAnalyzer();
            bool result = analyzer.HasDataRowsAndShape(sheet);

            Console.WriteLine($"Worksheet \"{sheet.Name}\" has data rows and shapes: {result}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
