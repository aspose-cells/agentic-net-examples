// Title: C# console program to load an Excel workbook and list each worksheet’s paper size and dimensions using Aspose.Cells
// AI Prompts: Create a C# console application that opens a given .xlsx file with Aspose.Cells, verifies the file exists, generates a placeholder workbook with a default sheet if it is missing, and then iterates through all worksheets to output the worksheet name, PageSetup.PaperSize, PageSetup.PaperWidth, and PageSetup.PaperHeight. | Add a try‑catch block around the workbook loading and processing logic that writes any caught exception messages to the console. | Ensure the placeholder workbook is saved to the original file path before the program proceeds to display the paper dimensions.
// Common Searches: how to read worksheet paper size with Aspose.Cells in a C# console app | C# Aspose.Cells print page setup dimensions for each Excel sheet | create empty Excel file if not found using Aspose.Cells .NET | retrieve PaperWidth and PaperHeight values from a worksheet with Aspose.Cells | Aspose.Cells console application list sheet paper size enum
// Tags: Aspose.Cells retrieve worksheet page setup dimensions | C# console load Excel workbook Aspose.Cells | Aspose.Cells handle missing workbook file | Aspose.Cells iterate worksheets print paper size | Aspose.Cells get PaperWidth PaperHeight values

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates a C# console app that checks for an input.xlsx file, creates a placeholder workbook when the file is absent, loads the workbook with Aspose.Cells, iterates through each worksheet, and prints the worksheet name along with its PaperSize enum, PaperWidth, and PaperHeight (in points). All operations are wrapped in exception handling to report errors gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file (adjust as needed)
            string filePath = "input.xlsx";

            try
            {
                // Ensure the file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    // Create an empty workbook as a placeholder
                    Workbook placeholder = new Workbook();
                    placeholder.Worksheets[0].Name = "Sheet1";
                    placeholder.Save(filePath);
                    Console.WriteLine($"Created empty workbook at {filePath}");
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the page setup of the current worksheet
                    PageSetup pageSetup = sheet.PageSetup;

                    // Retrieve paper size (enum) and custom dimensions (in points)
                    PaperSizeType paperSize = pageSetup.PaperSize;
                    double paperWidth = pageSetup.PaperWidth;   // width in points
                    double paperHeight = pageSetup.PaperHeight; // height in points

                    // Output the worksheet name and its paper dimensions
                    Console.WriteLine($"Worksheet: {sheet.Name}");
                    Console.WriteLine($"  Paper Size: {paperSize}");
                    Console.WriteLine($"  Width: {paperWidth} points, Height: {paperHeight} points");
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
