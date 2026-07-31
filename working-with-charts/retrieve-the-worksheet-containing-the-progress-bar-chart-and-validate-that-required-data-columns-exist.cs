// Title: Find the worksheet with a Progress Bar chart and verify required columns using Aspose.Cells for .NET
// Description: Loads an XLSX file with Aspose.Cells, scans all worksheets to locate the first Bar chart (used as a Progress Bar), identifies the hosting worksheet, checks that specified columns (e.g., A and B) contain at least one non‑empty cell within the used range, reports the results, and optionally saves the workbook.
// Keywords: Aspose.Cells C# chart detection | locate worksheet with chart Aspose | progress bar chart validation .NET | verify Excel column data presence | check chart source columns Aspose.Cells | load workbook with data validation | iterate worksheets and charts C# | Excel bar chart automation | batch workbook validation Aspose | chart data integrity check
// Common Searches: Aspose.Cells find worksheet that contains a specific chart | C# code to locate a progress bar chart in Excel | validate that columns A and B have data in the chart sheet using Aspose | how to detect bar chart type with Aspose.Cells .NET | check for empty source columns before saving workbook Aspose
// Developer Intent: Identify the worksheet that hosts a progress‑bar (Bar) chart and confirm that its required data columns are populated.
// Use Cases: Automatically locate the sheet that contains a progress bar chart to apply further formatting or calculations. | Ensure source data columns are not empty before generating reports that embed the chart. | Integrate chart‑data validation into batch processing pipelines for multiple Excel workbooks.
// AI Prompts: Write C# code with Aspose.Cells that finds a bar chart named 'Progress' and verifies that columns A‑C contain numeric values. | Provide an Aspose.Cells snippet that logs missing data in chart source columns and throws an exception if any required column is empty. | Show how to iterate all worksheets and charts to collect locations of progress bar charts and validate their data ranges in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX file with Aspose.Cells, scans all worksheets to locate the first Bar chart (used as a Progress Bar), identifies the hosting worksheet, checks that specified columns (e.g., A and B) contain at least one non‑empty cell within the used range, reports the results, and optionally saves the workbook.
class ProgressBarChartValidator
{
    static void Main()
    {
        // Load the workbook (replace "input.xlsx" with your file path)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CheckDataValid = true; // optional validation while loading
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Locate the Progress Bar chart and its containing worksheet
        Chart progressChart = null;
        Worksheet chartWorksheet = null;

        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                // Assuming a Progress Bar chart is a Bar chart; adjust the condition as needed
                if (ch.Type == ChartType.Bar)
                {
                    progressChart = ch;
                    chartWorksheet = ws;
                    break;
                }
            }
            if (progressChart != null) break;
        }

        if (progressChart == null)
        {
            Console.WriteLine("Progress Bar chart not found in the workbook.");
            return;
        }

        Console.WriteLine($"Progress Bar chart found in worksheet: {chartWorksheet.Name}");

        // Validate that required data columns exist (e.g., columns A and B must contain data)
        int[] requiredColumns = { 0, 1 }; // 0 = A, 1 = B

        // Determine the used row range of the worksheet
        int startRow = chartWorksheet.Cells.MinDataRow;
        int endRow   = chartWorksheet.Cells.MaxDataRow;

        foreach (int colIndex in requiredColumns)
        {
            bool hasData = false;

            for (int row = startRow; row <= endRow; row++)
            {
                Cell cell = chartWorksheet.Cells[row, colIndex];
                if (cell != null && cell.Type != CellValueType.IsNull && !string.IsNullOrEmpty(cell.StringValue))
                {
                    hasData = true;
                    break;
                }
            }

            Console.WriteLine($"Column {(char)('A' + colIndex)} contains data: {hasData}");
        }

        // Save the workbook after validation (optional)
        workbook.Save("output.xlsx");
    }
}
