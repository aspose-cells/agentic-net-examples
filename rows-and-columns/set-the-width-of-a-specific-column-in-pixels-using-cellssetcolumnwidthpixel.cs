// Title: How to set a worksheet column width in pixels with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Cells.SetColumnWidthPixel to set column C (index 2) to 150 px and saves the workbook. | Demonstrate adjusting multiple column widths in pixels using Aspose.Cells before populating data in a .NET workbook. | Show how to retrieve the Cells collection from a worksheet and apply SetColumnWidthPixel for precise column sizing in an Excel file.
// Common Searches: Aspose.Cells C# set column width by pixel example | How to use SetColumnWidthPixel to define exact column size in Excel with .NET | Set column C width to 150 pixels using Aspose.Cells API | Adjust Excel column widths in pixels programmatically with Aspose.Cells for C#
// Tags: Aspose.Cells SetColumnWidthPixel method | C# set Excel column width pixels | programmatic column width adjustment Aspose.Cells | Excel column pixel sizing .NET | worksheet column width precision Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, accesses the first worksheet, and uses Cells.SetColumnWidthPixel to set column index 2 (column C) to 150 pixels, writes a label into cell C1, saves the file as SetColumnWidthPixelDemo.xlsx, and includes basic exception handling.
    public class SetColumnWidthPixelDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                using (Workbook workbook = new Workbook())
                {
                    // Access the first worksheet
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Get the Cells collection
                    Cells cells = worksheet.Cells;

                    // Set the width of column 2 (third column, zero‑based index) to 150 pixels
                    cells.SetColumnWidthPixel(2, 150);

                    // Put some data to visualize the column width
                    cells["C1"].PutValue("Column C with 150px width");

                    // Save the workbook
                    string outputPath = "SetColumnWidthPixelDemo.xlsx";
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
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
