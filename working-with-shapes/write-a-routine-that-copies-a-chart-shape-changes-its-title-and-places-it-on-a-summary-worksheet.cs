// Title: Copy a ChartShape, rename its title, and place it on a summary sheet with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data and a column chart, retrieves the chart's ChartShape, copies it to a new "Summary" worksheet using ShapeCollection.AddCopy, changes the copied chart's title, and saves the file as ChartCopySummary.xlsx.
// Keywords: Aspose.Cells | .NET | ChartShape | AddCopy | copy chart | change chart title | summary worksheet | C# example
// Common Searches: Aspose.Cells copy chart to another worksheet | duplicate chart shape C# Aspose.Cells | set title of copied chart Aspose.Cells | ShapeCollection.AddCopy ChartShape example | copy and rename chart in Aspose.Cells
// Developer Intent: Duplicate an existing chart shape, modify its title, and store it on a summary worksheet.
// Use Cases: Build a dashboard that consolidates source charts onto a single summary page with uniform titles. | Generate a financial report that reuses charts from multiple sheets without recreating them. | Automate chart reuse across worksheets while applying custom titles for each copy.
// AI Prompts: Show C# code using Aspose.Cells to copy a ChartShape from one worksheet to another and set a new title. | Explain how ShapeCollection.AddCopy works for ChartShape objects and how to access the copied chart for property changes. | Provide a loop that copies several charts to a summary sheet and assigns distinct titles programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data and a column chart, retrieves the chart's ChartShape, copies it to a new "Summary" worksheet using ShapeCollection.AddCopy, changes the copied chart's title, and saves the file as ChartCopySummary.xlsx.
public class ChartCopyRoutine
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare source worksheet with a sample chart
            // -------------------------------------------------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Data";

            // Populate some data for the chart
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["A4"].PutValue("C");
            sourceSheet.Cells["B1"].PutValue("Value");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sourceSheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // 2. Access the ChartShape (the visual object of the chart)
            // -------------------------------------------------
            ChartShape sourceChartShape = chart.ChartObject; // ChartObject returns ChartShape

            // -------------------------------------------------
            // 3. Create (or get) a summary worksheet where the copy will be placed
            // -------------------------------------------------
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");

            // -------------------------------------------------
            // 4. Copy the chart shape to the summary worksheet
            //    Using ShapeCollection.AddCopy(sourceShape, topRow, top, leftColumn, left)
            // -------------------------------------------------
            ShapeCollection summaryShapes = summarySheet.Shapes;
            // Position the copied chart at row 2, column 2 (pixel offsets set to 0)
            Shape copiedShape = summaryShapes.AddCopy(sourceChartShape, 2, 0, 2, 0);

            // -------------------------------------------------
            // 5. Change the title of the copied chart shape
            //    The returned Shape is actually a ChartShape, so cast it
            // -------------------------------------------------
            if (copiedShape is ChartShape copiedChartShape)
            {
                copiedChartShape.Title = "Summary Chart";
            }

            // -------------------------------------------------
            // 6. Save the workbook
            // -------------------------------------------------
            workbook.Save("ChartCopySummary.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        ChartCopyRoutine.Run();
    }
}
