// Title: Create a Cross‑Sheet Line Sparkline Group in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to generate a workbook with a "DataSheet" worksheet containing sample data, add a "Sparklines" sheet, and place a line SparklineGroup in cells E1:E5 that references the range DataSheet!A1:D5. Shows adding an extra sparkline to a specific cell and saving the file as CrossSheetSparkline.xlsx.
// Keywords: Aspose.Cells | C# sparkline cross sheet | line sparkline group | SparklineGroup DataSheet range | Aspose.Cells example | cross‑worksheet sparkline | .NET workbook sparkline
// Common Searches: Aspose.Cells create sparkline from another worksheet | C# cross sheet sparkline example | How to add line sparkline group in Aspose.Cells | SparklineGroup with external sheet range .NET | Save workbook with cross‑sheet sparklines Aspose
// Developer Intent: Add a line sparkline group on one sheet that pulls its source data from a range on a different worksheet using Aspose.Cells for .NET.
// Use Cases: Display trend lines on a summary sheet while keeping raw data hidden on a separate DataSheet. | Build a sales or KPI dashboard where monthly values reside on a data sheet and sparklines visualize each item on a reporting sheet. | Programmatically insert additional sparklines into specific cells after the initial group is created for custom highlighting.
// AI Prompts: Generate C# code to create a SparklineGroup that references "DataSheet!B2:G10" and plots by column in Aspose.Cells. | Show how to customize color, weight, and marker style of a cross‑sheet sparkline group in Aspose.Cells for .NET. | Explain how to set the sparkline source range dynamically based on user‑selected cells on another worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineCrossSheetDemo
{
    // Demonstrates how to generate a workbook with a "DataSheet" worksheet containing sample data, add a "Sparklines" sheet, and place a line SparklineGroup in cells E1:E5 that references the range DataSheet!A1:D5. Shows adding an extra sparkline to a specific cell and saving the file as CrossSheetSparkline.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet that will hold the source data
                Worksheet dataSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                dataSheet.Name = "DataSheet";

                // Populate sample data in DataSheet (A1:D5)
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 4; col++)
                    {
                        dataSheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                    }
                }

                // Use the first worksheet to place the sparklines
                Worksheet sparklineSheet = workbook.Worksheets[0];
                sparklineSheet.Name = "Sparklines";

                // Define where the sparklines will be displayed (E1:E5) – one sparkline per row of data
                CellArea sparklineLocation = CellArea.CreateCellArea("E1", "E5");

                // Add a sparkline group that references the range on DataSheet
                // Data range format: SheetName!StartCell:EndCell
                int groupIndex = sparklineSheet.SparklineGroups.Add(
                    SparklineType.Line,               // Sparkline type
                    "DataSheet!A1:D5",                // Cross‑sheet data range
                    false,                            // Plot by row (horizontal)
                    sparklineLocation);               // Location range for sparklines

                // Retrieve the created group (optional, for further customization)
                SparklineGroup group = sparklineSheet.SparklineGroups[groupIndex];

                // Example: add another sparkline in cell E2 using the same data range
                group.Sparklines.Add("DataSheet!A1:D5", 1, 4); // row index 1 (E2), column index 4 (E)

                // Save the workbook
                workbook.Save("CrossSheetSparkline.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
