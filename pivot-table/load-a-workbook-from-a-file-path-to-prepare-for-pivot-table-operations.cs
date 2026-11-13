// Title: Load an Excel workbook from a file path with Aspose.Cells for .NET to enable pivot table creation
// Common Searches: C# Aspose.Cells load workbook from local file for pivot table generation | How to open an existing .xlsx file using Aspose.Cells before creating a pivot table | Aspose.Cells .NET read Excel workbook from path and access first worksheet | Initialize Aspose.Cells Workbook object from file path for pivot table manipulation | Sample code to load Excel file with Aspose.Cells in C# for pivot table operations
// Tags: Aspose.Cells workbook loading from file | C# open Excel file with Aspose.Cells | setup Aspose.Cells Workbook for pivot tables | read .xlsx using Aspose.Cells .NET | prepare worksheet for pivot table creation

using System;
using Aspose.Cells;

// Demonstrates loading an Excel workbook from a specified file path into an Aspose.Cells Workbook object in C#, accessing the first worksheet, and confirming successful loading, ready for subsequent pivot table operations.
class PivotTablePreparation
{
    static void Main()
    {
        // Specify the path to the Excel file.
        string filePath = @"C:\Data\SampleWorkbook.xlsx";

        // Load the workbook from the specified file.
        // This creates a Workbook object that can be used for further operations,
        // such as creating or manipulating pivot tables.
        Workbook workbook = new Workbook(filePath);

        // At this point the workbook is loaded and ready for pivot table operations.
        // Example: Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];

        // (Further pivot table code would go here.)

        // Optional: Keep the console open to verify successful loading.
        Console.WriteLine("Workbook loaded successfully. Worksheet count: " + workbook.Worksheets.Count);
    }
}
