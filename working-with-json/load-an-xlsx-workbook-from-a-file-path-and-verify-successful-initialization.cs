// Title: Load an XLSX workbook from a file path and confirm successful initialization with Aspose.Cells for .NET
// Common Searches: how to open an xlsx file using Aspose.Cells in C# | check if Aspose.Cells workbook loaded correctly in .NET | determine worksheet count after loading Excel file with Aspose.Cells | C# code to load Excel workbook from disk and verify it contains worksheets
// Tags: Aspose.Cells load workbook from file path | C# verify workbook initialization Aspose.Cells | worksheet count validation after workbook load | load xlsx workbook using Aspose.Cells .NET | check workbook null and worksheets collection Aspose.Cells

using System;
using Aspose.Cells;

// The example demonstrates creating a Workbook object from a specified XLSX file path with Aspose.Cells for .NET, then validates that the workbook instance is not null and contains at least one worksheet, outputting the worksheet count on success or an error message on failure.
class Program
{
    static void Main()
    {
        // Path to the XLSX file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Verify successful initialization
        if (workbook != null && workbook.Worksheets.Count > 0)
        {
            Console.WriteLine("Workbook loaded successfully. Number of worksheets: " + workbook.Worksheets.Count);
        }
        else
        {
            Console.WriteLine("Failed to load workbook or workbook contains no worksheets.");
        }
    }
}
