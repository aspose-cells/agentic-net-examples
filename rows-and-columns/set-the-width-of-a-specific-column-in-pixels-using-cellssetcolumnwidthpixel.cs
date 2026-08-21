// Title: Set column width in pixels with Cells.SetColumnWidthPixel – Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, access the first worksheet, and assign pixel‑based widths (e.g., 150 px for column C and 80 px for column A) using the Cells.SetColumnWidthPixel method in C# with Aspose.Cells, then save the workbook.
// Keywords: Aspose.Cells | Cells.SetColumnWidthPixel | set column width pixel | C# Aspose.Cells column width | Excel column width pixels | .NET column width | adjust column width Aspose | pixel based column width | Aspose.Cells column sizing
// Common Searches: Aspose.Cells set column width in pixels C# | Cells.SetColumnWidthPixel example | How to set column width to 150 pixels with Aspose.Cells | C# set first column width to 80 pixels Aspose | pixel based column width Aspose.Cells .NET
// Developer Intent: The developer wants to define exact pixel widths for specific worksheet columns using Aspose.Cells for .NET.
// Use Cases: Create a report where the first column needs a compact 80‑pixel width for IDs. | Allocate 150 pixels to a description column to prevent text wrapping. | Apply consistent pixel‑based column sizing across multiple worksheets before exporting to Excel.
// AI Prompts: Generate C# code that sets pixel widths for a range of columns using a loop with Cells.SetColumnWidthPixel. | Explain error‑handling strategies when applying pixel‑based column widths in Aspose.Cells. | Show how to convert point or character width measurements to pixel values for column sizing in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, access the first worksheet, and assign pixel‑based widths (e.g., 150 px for column C and 80 px for column A) using the Cells.SetColumnWidthPixel method in C# with Aspose.Cells, then save the workbook.
    public class SetColumnWidthPixelDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Set column widths in pixels
                cells.SetColumnWidthPixel(2, 150); // third column
                cells.SetColumnWidthPixel(0, 80);  // first column

                // Define output file path
                string outputPath = "ColumnWidthPixelDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetColumnWidthPixelDemo.Run();
        }
    }
}
