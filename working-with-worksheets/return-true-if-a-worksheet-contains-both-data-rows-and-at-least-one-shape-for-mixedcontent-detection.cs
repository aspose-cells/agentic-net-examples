// Title: Check if an Excel worksheet has any data rows and at least one shape with Aspose.Cells for .NET
// AI Prompts: Write a C# method that returns true only when a Worksheet contains at least one populated row and at least one drawing object, using Aspose.Cells. | Show how to load an Excel workbook with Aspose.Cells and iterate through its worksheets to flag those that have both data rows and shapes. | Generate code that evaluates Worksheet.Cells.MaxDataRow and Worksheet.Shapes.Count to determine mixed‑content presence in a sheet.
// Common Searches: Aspose.Cells C# check worksheet for both data rows and shapes | how to detect mixed content (data and drawings) in an Excel sheet using Aspose.Cells | C# find worksheets that contain at least one non‑empty row and a shape with Aspose.Cells | determine if Excel worksheet has any shapes and data rows using Aspose.Cells .NET
// Tags: worksheet data row detection Aspose.Cells | worksheet shape count Aspose.Cells | mixed content check Excel Aspose.Cells | maxdatarow shape presence C# | detect drawings in Excel worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example defines a static C# method that returns true when a given Worksheet has at least one data row (MaxDataRow >= 0) and at least one shape (Shapes.Count > 0). The program loads a workbook, selects the first worksheet, calls the method, and prints the result.
public static class WorksheetAnalyzer
{
    /// <param name="worksheet">The worksheet to inspect.</param>
    /// <returns>True if both data rows and shapes are present; otherwise, false.</returns>
    public static bool HasDataRowsAndShape(Worksheet worksheet)
    {
        // Check for data rows: MaxDataRow returns the zero‑based index of the last row that contains data.
        // If no data is present, MaxDataRow will be -1.
        bool hasDataRows = worksheet.Cells.MaxDataRow >= 0;

        // Check for shapes: the Shapes collection holds all drawing objects on the sheet.
        bool hasShapes = worksheet.Shapes.Count > 0;

        // Return true only when both conditions are satisfied.
        return hasDataRows && hasShapes;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Determine workbook path: use first argument or default sample file.
        string workbookPath = args.Length > 0 ? args[0] : "Sample.xlsx";

        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"File not found: {workbookPath}");
            return;
        }

        try
        {
            // Load workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Analyze worksheet.
            bool result = WorksheetAnalyzer.HasDataRowsAndShape(sheet);

            Console.WriteLine($"Worksheet \"{sheet.Name}\" has data rows and shapes: {result}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
