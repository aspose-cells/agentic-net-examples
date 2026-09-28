// Title: Load an XLSX file into an Aspose.Cells Workbook and confirm worksheet collection in C#
// AI Prompts: Generate C# code that opens a given .xlsx file with Aspose.Cells, checks that the Worksheets collection is instantiated, and prints the number of sheets. | Write a reusable function `bool IsWorkbookReady(string path)` that loads the workbook using Aspose.Cells and returns true only if at least one worksheet is present.
// Common Searches: C# Aspose.Cells how to open an .xlsx file and verify sheets exist | check if Excel workbook loaded with Aspose.Cells has any worksheets | Aspose.Cells .NET load workbook from file path and validate worksheet count | sample code to confirm successful workbook initialization using Aspose.Cells
// Tags: load xlsx workbook Aspose.Cells C# | verify worksheet collection Aspose.Cells | initialize workbook from file path .NET | check sheet count after loading Excel Aspose.Cells | Aspose.Cells workbook validation C#

using System;
using Aspose.Cells;

// The example demonstrates loading an .xlsx file from a specified path using Aspose.Cells for .NET, then confirming that the Workbook object and its Worksheets collection are instantiated and contain at least one sheet.
class Program
{
    static void Main()
    {
        // Path to the XLSX file
        string filePath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Verify successful initialization
        if (workbook != null && workbook.Worksheets != null && workbook.Worksheets.Count > 0)
        {
            Console.WriteLine("Workbook loaded successfully. Sheet count: " + workbook.Worksheets.Count);
        }
        else
        {
            Console.WriteLine("Failed to load workbook.");
        }
    }
}
