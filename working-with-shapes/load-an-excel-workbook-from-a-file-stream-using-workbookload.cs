// Title: Load an Excel workbook from a file path with Aspose.Cells Workbook constructor in C#
// AI Prompts: Generate C# code that opens a .xlsx file from a specified path using Aspose.Cells, verifies the file exists, loads the workbook, and prints the total number of worksheets. | Provide a C# example that demonstrates loading an Excel workbook with Aspose.Cells, includes try‑catch blocks for FileNotFoundException and other errors, and outputs the worksheet count.
// Common Searches: c# aspose.cells load workbook from file path with existence check | how to read worksheet count after loading Excel file using Aspose.Cells in C# | aspocells workbook constructor example handling file not found exception | load .xlsx into Aspose.Cells workbook from local file C# | exception handling when opening Excel workbook with Aspose.Cells
// Tags: open Excel file by path Aspose.Cells C# | Aspose.Cells Workbook constructor usage | validate Excel file existence before loading | handle workbook loading exceptions Aspose.Cells | retrieve worksheet count after loading workbook | C# read .xlsx with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks if "sample.xlsx" exists, loads the workbook using Aspose.Cells' Workbook constructor, prints the number of worksheets, and catches any exceptions that occur during loading.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "sample.xlsx";

        try
        {
            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook directly from the file path
            Workbook workbook = new Workbook(filePath);

            // Workbook is now loaded and can be used
            Console.WriteLine($"Workbook loaded successfully. Worksheets count: {workbook.Worksheets.Count}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
