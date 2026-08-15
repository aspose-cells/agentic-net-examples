// Title: C# Example: Set Uniform Row Height for All Rows with Aspose.Cells Cells.StandardHeight
// Description: Creates a new Workbook, accesses the first Worksheet, sets Cells.StandardHeight to a fixed value (e.g., 25 points) so every row shares the same height, adds sample data, and saves the file as UniformRowHeight.xlsx.
// Keywords: Aspose.Cells | C# row height | Cells.StandardHeight | uniform row height | set row height all rows | Excel worksheet row height | Aspose.Cells example | C# Excel automation | default row height | Aspose.Cells tutorial
// Common Searches: Aspose.Cells set same row height for entire worksheet C# | How to use Cells.StandardHeight in C# | C# code to apply uniform row height in Excel with Aspose.Cells | Set default row height for all rows Aspose.Cells
// Developer Intent: Apply a single row height to every row in a worksheet.
// Use Cases: Create a printable spreadsheet template with consistent row spacing. | Generate tabular reports where each row must have the same visual height. | Prepare data sheets before populating content to ensure uniform row layout. | Set a baseline row height before customizing specific rows for headers or footers.
// AI Prompts: Generate a reusable C# method that takes a height value and applies it to all rows using Cells.StandardHeight. | Show how to override the uniform height for selected rows after setting Cells.StandardHeight. | Provide an example that adjusts row height dynamically based on cell content length with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, accesses the first Worksheet, sets Cells.StandardHeight to a fixed value (e.g., 25 points) so every row shares the same height, adds sample data, and saves the file as UniformRowHeight.xlsx.
    public class UniformRowHeightDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Apply a uniform row height (e.g., 25 points) to all rows in the worksheet
                worksheet.Cells.StandardHeight = 25;

                // Add sample data to demonstrate the applied row height
                worksheet.Cells["A1"].PutValue("Row 1");
                worksheet.Cells["A2"].PutValue("Row 2");
                worksheet.Cells["A3"].PutValue("Row 3");

                // Define output file path
                string outputPath = "UniformRowHeight.xlsx";

                // Save the workbook to a file
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
            UniformRowHeightDemo.Run();
        }
    }
}
