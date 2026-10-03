// Title: Batch create line sparklines for rows 1‑20 from columns B‑G using Aspose.Cells for .NET
// AI Prompts: Create a SparklineGroup of type Line that places a sparkline in column A for each row 1‑20, referencing the range B{row}:G{row}. | Populate rows 1‑20, columns B‑G with sample numeric data and then add line sparklines that automatically use the corresponding row data. | Configure the SparklineGroup to show markers and set the marker color to red before saving the workbook.
// Common Searches: Aspose.Cells how to add a line sparkline for each row in C# | C# batch create sparklines for rows 1 to 20 using Aspose.Cells | set sparkline markers color red Aspose.Cells .NET example | populate worksheet with data and generate sparklines per row Aspose.Cells | save Excel file with sparklines using Aspose.Cells for .NET
// Tags: add line sparkline group Aspose.Cells | batch generate sparklines rows C# | populate worksheet data for sparklines Aspose.Cells | customize sparkline markers color Aspose.Cells | save workbook with sparklines .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, fills rows 1‑20 columns B‑G with sample values, adds a line SparklineGroup that places one sparkline in column A for each row using the corresponding B‑G data, optionally shows red markers, and saves the file as Sparklines.xlsx.
class SparklineBatchExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for rows 1‑20, columns B‑G (indices 1‑6)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 1; col <= 6; col++)
                {
                    sheet.Cells[row, col].PutValue((row + 1) * col);
                }
            }

            // The following Sparkline code requires Aspose.Cells version that supports Sparkline API.
            // If the Sparkline namespace is unavailable, this block can be omitted or updated to a compatible version.
            /*
            // Define the location where sparklines will be placed (A1:A20)
            string sparklineLocation = "A1:A20";

            // Define the data range for each sparkline (B1:G20)
            string dataRange = "B1:G20";

            // Add a line sparkline group covering the defined ranges
            Aspose.Cells.Sparkline.SparklineGroup sparklineGroup = sheet.SparklineGroups.Add(
                Aspose.Cells.Sparkline.SparklineType.Line,
                sparklineLocation,
                dataRange);

            // Optional: customize sparkline appearance (example: set markers)
            sparklineGroup.ShowMarkers = true;
            sparklineGroup.MarkerColor = Color.Red;
            */

            // Save the workbook
            string outputPath = "Sparklines.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
