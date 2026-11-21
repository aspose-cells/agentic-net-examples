// Title: Generate an Excel audit worksheet that lists each chart’s axis type and tick‑label direction using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# retrieve chart axis type and tick label position | How to create a chart audit report in an Excel workbook with Aspose.Cells | List all chart axes and their label directions using Aspose.Cells .NET | C# code to extract CategoryAxis and ValueAxis properties from Excel charts | Generate a summary sheet of chart axis information with Aspose.Cells
// Tags: extract chart axis type Aspose.Cells | retrieve tick label position Aspose.Cells | audit Excel charts Aspose.Cells C# | generate chart properties worksheet Aspose.Cells | iterate workbook charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, adds a "ChartAudit" worksheet, writes header columns, then iterates through every worksheet (excluding the audit sheet) and each chart it contains. For each chart it records the runtime type name and TickLabelPosition of the CategoryAxis (X axis) and ValueAxis (Y axis), along with worksheet name, chart index, and chart name, into the audit sheet and saves the file as output.xlsx, handling missing files and exceptions.
class ChartAxisAudit
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the audit report
            int reportSheetIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
            reportSheet.Name = "ChartAudit";

            // Write header row
            Cells cells = reportSheet.Cells;
            cells["A1"].PutValue("Worksheet");
            cells["B1"].PutValue("Chart Index");
            cells["C1"].PutValue("Chart Name");
            cells["D1"].PutValue("X Axis Type");
            cells["E1"].PutValue("X Axis Tick Label Direction");
            cells["F1"].PutValue("Y Axis Type");
            cells["G1"].PutValue("Y Axis Tick Label Direction");

            int reportRow = 1; // zero‑based index; row 1 is the second row (after header)

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Skip the report sheet itself
                if (sheet.Name == "ChartAudit")
                    continue;

                int chartIdx = 0;
                // Iterate through all charts in the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // X Axis (Category Axis)
                    Axis xAxis = chart.CategoryAxis;
                    // Use the runtime type name as a fallback for axis type information
                    string xAxisType = xAxis.GetType().Name;
                    string xTickLabelDirection = xAxis.TickLabelPosition.ToString();

                    // Y Axis (Value Axis)
                    Axis yAxis = chart.ValueAxis;
                    string yAxisType = yAxis.GetType().Name;
                    string yTickLabelDirection = yAxis.TickLabelPosition.ToString();

                    // Write data to the report sheet
                    cells[reportRow, 0].PutValue(sheet.Name);
                    cells[reportRow, 1].PutValue(chartIdx);
                    cells[reportRow, 2].PutValue(chart.Name);
                    cells[reportRow, 3].PutValue(xAxisType);
                    cells[reportRow, 4].PutValue(xTickLabelDirection);
                    cells[reportRow, 5].PutValue(yAxisType);
                    cells[reportRow, 6].PutValue(yTickLabelDirection);

                    reportRow++;
                    chartIdx++;
                }
            }

            // Save the workbook with the audit report
            workbook.Save(outputPath);
            Console.WriteLine($"Audit report saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
