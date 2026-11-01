// Title: Generate a chart axis audit worksheet that lists each axis type and tick‑label position using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to iterate all worksheets, extract each chart’s CategoryAxis and ValueAxis properties, and record the axis class name and TickLabelPosition into a new sheet. | Modify the sample to also include secondary axes and output the collected axis information to a CSV file.
// Common Searches: how to list axis type and tick label position for all charts in an Excel file using Aspose.Cells C# | Aspose.Cells C# create report of chart axes properties | retrieve chart axis TickLabelPosition with Aspose.Cells .NET | export chart axis metadata to a new worksheet in Aspose.Cells | C# enumerate charts and write axis details to a summary sheet
// Tags: Aspose.Cells extract chart axis properties | C# generate chart axis audit worksheet | Aspose.Cells enumerate workbook charts | record TickLabelPosition Aspose.Cells | write axis metadata to Excel sheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, adds a worksheet named "Chart Axis Report", walks through every worksheet and its charts, captures the Category and Value axis types (using the runtime class name) and their TickLabelPosition values, writes these details to the report sheet, auto‑fits columns, and saves the workbook with the new audit data.
class ChartAxisReport
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the audit report
            int reportSheetIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
            reportSheet.Name = "Chart Axis Report";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("Chart Index");
            reportSheet.Cells[0, 2].PutValue("Axis");
            reportSheet.Cells[0, 3].PutValue("Axis Type");
            reportSheet.Cells[0, 4].PutValue("Tick Label Position");

            int reportRow = 1;

            // Iterate through all worksheets except the report sheet itself
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == reportSheet.Name) continue;

                // Process each chart on the worksheet
                for (int i = 0; i < ws.Charts.Count; i++)
                {
                    Chart chart = ws.Charts[i];

                    // Primary Category (X) Axis
                    Axis categoryAxis = chart.CategoryAxis;
                    if (categoryAxis != null)
                    {
                        WriteAxisInfo(reportSheet, ref reportRow, ws.Name, i, "Category (X)", categoryAxis);
                    }

                    // Primary Value (Y) Axis
                    Axis valueAxis = chart.ValueAxis;
                    if (valueAxis != null)
                    {
                        WriteAxisInfo(reportSheet, ref reportRow, ws.Name, i, "Value (Y)", valueAxis);
                    }

                    // Note: Secondary axes are not accessed here due to API version differences.
                }
            }

            // Adjust column widths for readability
            reportSheet.AutoFitColumns();

            // Save the workbook with the added report
            string outputPath = "output_with_chart_axis_report.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Report saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to write axis information to the report sheet
    private static void WriteAxisInfo(Worksheet sheet, ref int row, string worksheetName, int chartIndex, string axisLabel, Axis axis)
    {
        sheet.Cells[row, 0].PutValue(worksheetName);
        sheet.Cells[row, 1].PutValue(chartIndex);
        sheet.Cells[row, 2].PutValue(axisLabel);
        // Use the runtime type name as a fallback for axis type information
        sheet.Cells[row, 3].PutValue(axis.GetType().Name);
        // Use TickLabelPosition as a representative property for label direction/position
        sheet.Cells[row, 4].PutValue(axis.TickLabelPosition.ToString());

        row++;
    }
}
