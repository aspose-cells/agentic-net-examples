// Title: C# – Insert a Win/Loss Sparkline in cell E2 using Aspose.Cells for .NET
// Description: Creates a new workbook, populates F2:F12 with sample values, places a stacked win/loss sparkline in E2, and saves the file as WinLossSparkline.xlsx.
// Keywords: Aspose.Cells win loss sparkline C# | add sparkline to Excel .NET | stacked sparkline group Aspose | sparkline cell E2 example | populate sparkline data range F2:F12
// Common Searches: how to add a win/loss sparkline with Aspose.Cells | C# code for sparkline in cell E2 from range F2:F12 | Aspose.Cells example stacked sparkline
// Developer Intent: Generate a win/loss (stacked) sparkline in E2 that visualizes the numeric series in F2:F12.
// Use Cases: Show daily profit/loss trends in a compact financial report. | Add visual indicators to a KPI dashboard for each product row. | Automate sparkline creation across multiple rows in a budgeting worksheet.
// AI Prompts: Generate C# Aspose.Cells code to place a win/loss sparkline in D5 based on data in G2:G15. | Explain how to enable positive/negative markers for a stacked sparkline group in Aspose.Cells. | Write a method that loops rows 2‑10 and inserts a win/loss sparkline in column E for each corresponding column F range.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, populates F2:F12 with sample values, places a stacked win/loss sparkline in E2, and saves the file as WinLossSparkline.xlsx.
class InsertWinLossSparkline
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (F2:F12)
            // Positive values represent gains, negative values represent losses
            double[] sampleData = { 10, -5, 8, -3, 12, -7, 4, -2, 6, -1, 9 };
            for (int i = 0; i < sampleData.Length; i++)
            {
                // Row index starts at 1 for the second row (F2)
                sheet.Cells[i + 1, 5].PutValue(sampleData[i]); // Column index 5 = column F
            }

            // Define the location where the sparkline will be placed (cell E2)
            CellArea location = new CellArea
            {
                StartColumn = 4, // Column E
                EndColumn = 4,
                StartRow = 1,    // Row 2 (zero‑based index)
                EndRow = 1
            };

            // Add a Win/Loss (Stacked) sparkline group.
            // Data range is F2:F12, plotted by column (isVertical = false)
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Stacked,   // Win/Loss sparkline type
                "F2:F12",                // Data range
                false,                   // Plot by column
                location);               // Location range (E2)

            // Save the workbook
            string outputPath = "WinLossSparkline.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
