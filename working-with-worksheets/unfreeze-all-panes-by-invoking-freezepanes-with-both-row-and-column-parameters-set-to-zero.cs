// Title: How to unfreeze all panes in an Excel worksheet with Aspose.Cells for .NET (C#) using FreezePanes(0,0,0,0)
// Common Searches: Aspose.Cells C# unfreeze worksheet panes programmatically | remove frozen rows and columns in Excel using FreezePanes zero parameters | C# code to clear pane freeze in Aspose.Cells workbook | how to reset frozen panes in Aspose.Cells .NET example | unfreeze all panes in an Excel file with Aspose.Cells API
// Tags: Aspose.Cells FreezePanes unfreeze panes | C# unfreeze Excel worksheet panes | Aspose.Cells remove frozen rows columns | Excel workbook pane reset Aspose .NET | FreezePanes zero parameters example

using System;
using System.IO;
using Aspose.Cells;

// This example creates a new Workbook (or loads an existing one), accesses the first Worksheet, and calls FreezePanes(0,0,0,0) to remove any frozen rows or columns. It ensures the output directory exists, saves the file as 'Unfrozen.xlsx', and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Unfreeze all panes. Use the 4‑parameter overload with zeros.
            sheet.FreezePanes(0, 0, 0, 0);

            // Define output file path
            string outputPath = "Unfrozen.xlsx";

            // Ensure the output directory exists (handle possible null)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
