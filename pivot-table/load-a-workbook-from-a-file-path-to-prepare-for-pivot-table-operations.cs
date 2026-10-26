// Title: Load an Excel workbook from a file path with Aspose.Cells for .NET to enable pivot table creation
// Common Searches: Aspose.Cells C# load workbook from local file for pivot table | how to open an existing .xlsx file using Aspose.Cells in .NET | C# example reading Excel workbook by path before creating pivot tables | using Aspose.Cells to initialize workbook for pivot table operations | load Excel file into Aspose.Cells Workbook object from file system
// Tags: Aspose.Cells load workbook from file path | C# Aspose.Cells open .xlsx workbook | initialize workbook for pivot tables Aspose.Cells | read Excel file with Aspose.Cells .NET | prepare worksheet for pivot table creation Aspose.Cells

using System;
using Aspose.Cells;

// The example demonstrates creating an Aspose.Cells Workbook by loading a .xlsx file from a specified path, accessing the first worksheet, and confirming the workbook is ready for further pivot table manipulation.
class Program
{
    static void Main()
    {
        // Specify the path to the Excel file
        string filePath = @"C:\Data\SampleWorkbook.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(filePath);

        // At this point the workbook is ready for pivot table operations
        // Example: access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // (Further pivot table code would go here)

        // Optional: keep the console open
        Console.WriteLine("Workbook loaded successfully.");
    }
}
