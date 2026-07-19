// Title: C# – Validate a Progress‑Bar Chart Worksheet and Required Columns with Aspose.Cells
// Description: Loads an Excel workbook (or creates a minimal fallback), scans every worksheet for a Bar or Column chart whose name contains "Progress", reports the chart location, and verifies that the header row includes the required "Task" and "Progress" columns before optionally saving the file. Demonstrates chart detection, header validation, and safe workbook handling using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | progress bar chart | validate worksheet columns | find chart by name | load workbook without validation | Excel header check | chart detection .NET | Excel automation | Aspose.Cells example
// Common Searches: How to locate a progress bar chart in Excel using Aspose.Cells C# | Validate required columns for a chart with Aspose.Cells | Find Bar or Column chart named Progress in a workbook | Load Excel file with LoadOptions that disables data validation | Check if Task and Progress headers exist in a worksheet
// Developer Intent: Identify the worksheet that contains a progress‑bar chart and confirm that the required "Task" and "Progress" columns are present.
// Use Cases: Automated quality‑check of Excel templates before publishing reports. | Dynamic validation of user‑provided workbooks in a web or desktop application. | Generating a fallback workbook with sample data when the expected template is missing. | Logging chart locations and missing headers for downstream processing or error reporting.
// AI Prompts: Write C# code with Aspose.Cells that iterates all worksheets, finds Bar or Column charts whose Name includes "Progress", and returns the worksheet name and index. | Create a method that scans the first row of a given worksheet for a list of header strings and logs which headers are missing, using Aspose.Cells APIs. | Show how to load an Excel file with LoadOptions that disables data validation, create a minimal workbook with "Task" and "Progress" headers if the file does not exist, and then save the result.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook (or creates a minimal fallback), scans every worksheet for a Bar or Column chart whose name contains "Progress", reports the chart location, and verifies that the header row includes the required "Task" and "Progress" columns before optionally saving the file. Demonstrates chart detection, header validation, and safe workbook handling using Aspose.Cells for .NET.
class ProgressBarChartValidator
{
    static void Main()
    {
        try
        {
            // Path to the template workbook
            string templatePath = "ProgressBarTemplate.xlsx";

            // Ensure the file exists; if not, create a minimal workbook as fallback
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                // Load the workbook (disable template validation if needed)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    CheckDataValid = false
                };
                workbook = new Workbook(templatePath, loadOptions);
            }
            else
            {
                // Create a new workbook with required headers for demonstration
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Data";
                ws.Cells[0, 0].PutValue("Task");
                ws.Cells[0, 1].PutValue("Progress");
                // Optionally add a simple column chart to mimic a progress‑bar chart
                int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                Chart chart = ws.Charts[chartIndex];
                chart.Name = "ProgressChart";
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";
            }

            // Columns that must be present for the chart data
            string[] requiredHeaders = { "Task", "Progress" };

            bool chartFound = false;

            // Search each worksheet for a progress‑bar style chart
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart chart in ws.Charts)
                {
                    // Treat Bar or Column charts as potential progress bar charts
                    if (chart.Type == ChartType.Bar || chart.Type == ChartType.Column)
                    {
                        // Optional: further identify by chart name containing "Progress"
                        if (!string.IsNullOrEmpty(chart.Name) &&
                            chart.Name.IndexOf("Progress", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            chartFound = true;
                            Console.WriteLine($"Progress bar chart found in worksheet '{ws.Name}' (index {ws.Index}).");
                            ValidateRequiredColumns(ws, requiredHeaders);
                        }
                    }
                }
            }

            if (!chartFound)
            {
                Console.WriteLine("No progress bar chart was found in the workbook.");
            }

            // Save the workbook after validation (optional)
            string outputPath = "ProgressBarValidated.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static void ValidateRequiredColumns(Worksheet ws, string[] headers)
    {
        // Assume header row is the first row (row index 0)
        int headerRow = 0;
        foreach (string header in headers)
        {
            bool exists = false;
            // Scan all columns in the header row
            for (int col = 0; col <= ws.Cells.MaxColumn; col++)
            {
                var cell = ws.Cells[headerRow, col];
                if (cell != null && cell.Type == CellValueType.IsString &&
                    cell.StringValue.Equals(header, StringComparison.OrdinalIgnoreCase))
                {
                    exists = true;
                    break;
                }
            }

            Console.WriteLine(exists
                ? $"Required column '{header}' exists."
                : $"Required column '{header}' is missing.");
        }
    }
}
