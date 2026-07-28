// Title: Hide Secondary Y‑Axis Tick Marks in a Mixed Column‑Line Chart – C# Aspose.Cells Example
// Description: This C# sample creates a workbook with category, sales, and profit data, adds a mixed column‑line chart, assigns the profit series to the secondary Y‑axis, and suppresses both major and minor tick marks on that axis by using TickMarkType.None. The chart is saved as an XLSX file for clean visual presentation.
// Keywords: Aspose.Cells | C# | mixed chart | secondary Y axis | hide tick marks | TickMarkType.None | chart axis formatting | column line chart | visual clutter reduction | Excel automation | Aspose.Cells for .NET
// Common Searches: Aspose.Cells hide secondary axis tick marks | C# remove tick marks from secondary Y axis | mixed column line chart Aspose.Cells | set secondary value axis tick mark none Aspose | how to hide minor tick marks in Aspose chart
// Developer Intent: Remove major and minor tick marks from the secondary Y‑axis of a mixed column‑line chart.
// Use Cases: Financial dashboards where profit percentages are shown on a secondary axis without distracting tick marks. | Automated report generation that requires a clean mixed chart for presentation slides. | Data‑driven Excel workbooks that need aesthetic axis formatting for client‑facing documents.
// AI Prompts: Show C# code to hide major and minor tick marks on the secondary Y‑axis of a mixed chart using Aspose.Cells. | How can I set TickMarkType.None for the secondary value axis in an Aspose.Cells column‑line chart? | Explain the steps to create a mixed chart and suppress secondary axis tick marks with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# sample creates a workbook with category, sales, and profit data, adds a mixed column‑line chart, assigns the profit series to the secondary Y‑axis, and suppresses both major and minor tick marks on that axis by using TickMarkType.None. The chart is saved as an XLSX file for clean visual presentation.
    class HideSecondaryYAxisTickMarks
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – categories
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["A5"].PutValue("Q4");

            // Column B – primary series (e.g., Sales)
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["B4"].PutValue(180);
            worksheet.Cells["B5"].PutValue(200);

            // Column C – secondary series (e.g., Profit Margin %)
            worksheet.Cells["C1"].PutValue("Profit %");
            worksheet.Cells["C2"].PutValue(15);
            worksheet.Cells["C3"].PutValue(18);
            worksheet.Cells["C4"].PutValue(20);
            worksheet.Cells["C5"].PutValue(22);

            // Add a mixed chart (Column for primary series, Line for secondary series)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Add primary series (column) and bind data
            chart.NSeries.Add("B2:B5", true);
            // Add secondary series (line) and bind data
            chart.NSeries.Add("C2:C5", true);
            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // Plot the second series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Hide tick marks on the secondary Y axis
            Axis secondaryValueAxis = chart.SecondValueAxis;
            secondaryValueAxis.MajorTickMark = TickMarkType.None; // Hide major tick marks
            secondaryValueAxis.MinorTickMark = TickMarkType.None; // Hide minor tick marks

            // Optional: adjust secondary axis title for clarity
            secondaryValueAxis.Title.Text = "Profit %";

            // Save the workbook
            workbook.Save("MixedChart_SecondaryYAxis_TickMarksHidden.xlsx");
        }
    }
}
