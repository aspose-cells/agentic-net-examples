using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ProgressBarChartValidator
{
    static void Main()
    {
        const string templatePath = "ProgressBarTemplate.xlsx";
        const string outputPath = "ProgressBarValidated.xlsx";

        try
        {
            // Verify that the template file exists before attempting to load it
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {Path.GetFullPath(templatePath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(templatePath);

            Worksheet chartWorksheet = null;
            Chart progressChart = null;

            // Locate the worksheet that contains the Progress Bar chart
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart chart in ws.Charts)
                {
                    // Identify the chart by its type (Bar or Column)
                    if (chart.Type == ChartType.Bar || chart.Type == ChartType.Column)
                    {
                        progressChart = chart;
                        chartWorksheet = ws;
                        break;
                    }
                }
                if (chartWorksheet != null) break;
            }

            if (chartWorksheet == null)
            {
                Console.WriteLine("Progress Bar chart not found.");
                return;
            }

            Console.WriteLine($"Progress Bar chart found in worksheet: {chartWorksheet.Name}");

            // Validate that required data columns exist (e.g., columns A and B)
            bool columnAHasData = false;
            bool columnBHasData = false;

            // Determine the last row with data in the worksheet
            int lastRow = chartWorksheet.Cells.MaxDataRow;

            // Check column A (index 0)
            for (int row = 0; row <= lastRow; row++)
            {
                Cell cell = chartWorksheet.Cells[row, 0];
                if (cell != null && cell.Type != CellValueType.IsNull && !string.IsNullOrWhiteSpace(cell.StringValue))
                {
                    columnAHasData = true;
                    break;
                }
            }

            // Check column B (index 1)
            for (int row = 0; row <= lastRow; row++)
            {
                Cell cell = chartWorksheet.Cells[row, 1];
                if (cell != null && cell.Type != CellValueType.IsNull && !string.IsNullOrWhiteSpace(cell.StringValue))
                {
                    columnBHasData = true;
                    break;
                }
            }

            // Report validation results
            if (columnAHasData && columnBHasData)
            {
                Console.WriteLine("Required data columns (A and B) are present.");
            }
            else
            {
                Console.WriteLine("Missing required data columns:");
                if (!columnAHasData) Console.WriteLine("- Column A");
                if (!columnBHasData) Console.WriteLine("- Column B");
            }

            // Save the workbook (optional – can be omitted if no changes are made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}