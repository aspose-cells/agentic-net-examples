// Title: Hide Column B in Aspose.Cells (C#) Using the EntireColumn Property
// Description: Demonstrates how to create a workbook, obtain the EntireColumn for range B1, retrieve its zero‑based index, hide column B with HideColumn, and save the file as HideColumnB.xlsx.
// Keywords: Aspose.Cells hide column C# | EntireColumn property | HideColumn method | column B Excel Aspose | .NET Excel column visibility
// Common Searches: Aspose.Cells hide specific column | C# hide column B using EntireColumn | How to hide a column in Aspose.Cells workbook | Retrieve column index from range Aspose.Cells
// Developer Intent: Programmatically hide column B by extracting its index via the EntireColumn property and calling HideColumn.
// Use Cases: Mask confidential data before sharing a workbook | Apply user‑defined column visibility settings | Temporarily collapse columns during dynamic report generation
// AI Prompts: Write C# code to hide multiple columns using the EntireColumn property in Aspose.Cells. | Show how to toggle column visibility with a boolean flag in Aspose.Cells for .NET. | Explain the steps to get a column index from a Range and hide that column using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Example demonstrating how to hide a column using the EntireColumn property.
    // Demonstrates how to create a workbook, obtain the EntireColumn for range B1, retrieve its zero‑based index, hide column B with HideColumn, and save the file as HideColumnB.xlsx.
    public class HideColumnUsingEntireColumn
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a range that starts at cell B1.
                // Use CreateRange to obtain an Aspose.Cells.Range object.
                Aspose.Cells.Range range = worksheet.Cells.CreateRange("B1");

                // Get the entire column that contains the range (column B).
                Aspose.Cells.Range entireColumn = range.EntireColumn;

                // Determine the zero‑based column index of the entire column.
                int columnIndex = entireColumn.FirstColumn; // Column B => index 1

                // Hide the column using the column index.
                worksheet.Cells.HideColumn(columnIndex);

                // Save the workbook.
                string outputPath = "HideColumnB.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application.
    internal class Program
    {
        private static void Main(string[] args)
        {
            HideColumnUsingEntireColumn.Run();
        }
    }
}
