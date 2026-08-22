// Title: Load an Excel workbook from a file path with Aspose.Cells for .NET to prepare for pivot table creation
// AI Prompts: Generate C# code that reads a specified .xlsx file using Aspose.Cells and returns a Workbook instance. | Write a console program that opens an existing Excel workbook with Aspose.Cells, then prints the number of worksheets to confirm the load. | Show how to instantiate an Aspose.Cells Workbook from a file path and verify successful loading before manipulating pivot tables.
// Common Searches: how to open an existing .xlsx file with Aspose.Cells in a C# console app | Aspose.Cells load workbook from local path for pivot table generation | C# example to read Excel file and get worksheet count using Aspose.Cells | initialize Aspose.Cells workbook before creating pivot tables in .NET
// Tags: Aspose.Cells retrieve workbook from file | C# prepare workbook for pivot table creation | access Excel worksheets count with Aspose.Cells | validate workbook loading Aspose.Cells

using System;
using Aspose.Cells;

// // Loads "input.xlsx" into an Aspose.Cells Workbook object and prints the total worksheet count, establishing a base for subsequent pivot table operations.
class LoadWorkbookForPivot
{
    static void Main()
    {
        // Path to the existing Excel file
        string filePath = "input.xlsx";

        // Load the workbook from the specified file path
        Workbook workbook = new Workbook(filePath);

        // The workbook is now ready for pivot table operations
        // Example: display basic information
        Console.WriteLine($"Workbook loaded. Worksheets count: {workbook.Worksheets.Count}");
    }
}
