// Title: Create a win/loss column sparkline in cell E2 from range F2:F12 using Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a win/loss column sparkline at E2 referencing the data range F2:F12 with Aspose.Cells. | Write a script to add a column sparkline group for a win/loss chart in an Excel workbook, placing it in cell E2 using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells how to add a win/loss sparkline to a single cell in C# | C# example for creating a column sparkline from F2 to F12 with Aspose.Cells | Insert a win/loss sparkline at E2 in an .xlsx file using Aspose.Cells for .NET | Programmatically add a sparkline group for win/loss data in Excel with Aspose.Cells
// Tags: win/loss column sparkline Aspose.Cells .NET | add sparkline to cell E2 Aspose.Cells | sparkline group creation C# Aspose.Cells | Excel sparkline generation .NET | save workbook with sparkline Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Sparkline types

namespace SparklineExample
{
    // The example creates a new workbook, adds a win/loss column sparkline that visualizes values in F2:F12, places it in cell E2, and saves the file as SparklineOutput.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define data range and location for the sparkline
                string dataRange = "F2:F12";
                // Sparkline location is a single cell (E2)
                CellArea location = CellArea.CreateCellArea("E2", "E2");

                // Add a Column sparkline group; Add returns the index of the created group
                int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Column, dataRange, false, location);
                SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

                // (Optional) Configure sparkline appearance here using sparklineGroup if needed

                // Define output file path
                string outputPath = "SparklineOutput.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
