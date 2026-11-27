// Title: Load an Excel workbook from a file path and retrieve the first worksheet using Aspose.Cells for .NET
// Common Searches: how to open an existing Excel file with Aspose.Cells in C# and get the first sheet | Aspose.Cells C# load workbook from specific path and read cell A1 | C# code to create a workbook if file not found then open it with Aspose.Cells | using Aspose.Cells to access worksheet index 0 after loading a workbook | example of Workbook(string) constructor usage in Aspose.Cells for .NET
// Tags: Aspose.Cells load workbook from file | first worksheet access Aspose.Cells C# | Workbook(string) usage Aspose.Cells | auto-create workbook when missing Aspose.Cells | read cell A1 Aspose.Cells worksheet

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an Excel file from a given path (or creating a minimal workbook if the file is absent) with Aspose.Cells for .NET, accessing the first worksheet, and reading the value of cell A1.
class Program
{
    static void Main(string[] args)
    {
        // Determine the Excel file path. Use the first command‑line argument if provided,
        // otherwise fall back to a default file name in the current directory.
        string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

        try
        {
            // If the file does not exist, create a minimal workbook so the load operation can succeed.
            if (!File.Exists(filePath))
            {
                Workbook creator = new Workbook();                     // Workbook()
                Worksheet sheet = creator.Worksheets[0];               // Access default worksheet
                sheet.Name = "Sheet1";
                sheet.Cells["A1"].PutValue("Created because file was missing");
                creator.Save(filePath, SaveFormat.Xlsx);               // Save(string, SaveFormat)
                Console.WriteLine($"Sample workbook created at '{Path.GetFullPath(filePath)}'.");
            }

            // Load the workbook from the specified path.
            Workbook workbook = new Workbook(filePath);                 // Workbook(string)

            // Access the first worksheet in the workbook.
            Worksheet firstWorksheet = workbook.Worksheets[0];         // WorksheetCollection indexer

            Console.WriteLine($"Workbook loaded from: {Path.GetFullPath(filePath)}");
            Console.WriteLine($"First worksheet name: {firstWorksheet.Name}");

            // Demonstrate reading a cell value from the first worksheet.
            Cell cellA1 = firstWorksheet.Cells["A1"];                  // Cells["A1"]
            Console.WriteLine($"Cell A1 value: {cellA1.StringValue}");
        }
        catch (Exception ex)
        {
            // Report any errors without suppressing them.
            Console.Error.WriteLine($"Error: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
