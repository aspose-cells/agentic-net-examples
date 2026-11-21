// Title: C# console application with Aspose.Cells to load an Excel workbook and output each worksheet’s paper size enum and dimensions
// Common Searches: how to read worksheet paper size and dimensions using Aspose.Cells in a C# console app | C# code to list Excel sheet page setup properties with Aspose.Cells | retrieve paper width and height in points for each worksheet via Aspose.Cells | command‑line tool for displaying worksheet paper size enum Aspose.Cells C# | Aspose.Cells example for iterating worksheets and printing page setup details
// Tags: Aspose.Cells worksheet page setup retrieval | Aspose.Cells get paper size enumeration | Aspose.Cells extract paper width height points | C# load Excel workbook iterate worksheets | Aspose.Cells console application example | Aspose.Cells display worksheet dimensions

using System;
using Aspose.Cells;

// A C# console program that accepts a workbook path, loads the Excel file with Aspose.Cells, iterates through all worksheets, and writes each sheet's name, PaperSize enum value, PaperWidth, and PaperHeight (in points) to the console.
class Program
{
    static void Main(string[] args)
    {
        // Verify that a workbook path was provided.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: dotnet run <workbook_path>");
            return;
        }

        string workbookPath = args[0];

        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(workbookPath);

        // Iterate through each worksheet in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the page setup for the current worksheet.
            PageSetup pageSetup = sheet.PageSetup;

            // Print worksheet name.
            Console.WriteLine($"Worksheet: {sheet.Name}");

            // Print the paper size enumeration (e.g., A4, Letter).
            Console.WriteLine($"Paper Size Enum: {pageSetup.PaperSize}");

            // Print the actual width and height in points (1 point = 1/72 inch).
            Console.WriteLine($"Paper Width: {pageSetup.PaperWidth} points");
            Console.WriteLine($"Paper Height: {pageSetup.PaperHeight} points");
            Console.WriteLine();
        }
    }
}
