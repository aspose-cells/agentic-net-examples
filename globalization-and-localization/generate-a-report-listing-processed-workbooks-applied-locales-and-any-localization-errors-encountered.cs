using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Define a simple class to hold report data for each workbook
class WorkbookReportInfo
{
    public string Path { get; set; }
    public string Locale { get; set; }
    public List<string> Errors { get; set; } = new List<string>();
}

// Custom globalization settings that store a locale identifier
class LocaleGlobalizationSettings : GlobalizationSettings
{
    private readonly string _locale;

    public LocaleGlobalizationSettings(string locale)
    {
        _locale = locale;
    }

    // Expose the locale for reporting purposes
    public string Locale => _locale;

    // You can override other methods if needed; default behavior is sufficient for this demo
}

// Main processing class
class LocalizationReportGenerator
{
    // Entry point
    public static void Main()
    {
        // List of workbook files to process
        List<string> workbookFiles = new List<string>
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx"
            // Add more file paths as needed
        };

        // Collect report information for each workbook
        List<WorkbookReportInfo> reports = new List<WorkbookReportInfo>();

        foreach (string filePath in workbookFiles)
        {
            // Load the workbook using the provided load rule
            Workbook wb = new Workbook(filePath);

            // Determine the locale to apply (for demonstration we use file name convention)
            // In a real scenario, the locale might be stored in a custom property or external config
            string locale = filePath.Contains("zh") ? "zh-CN" : "en-US";

            // Apply custom globalization settings
            wb.Settings.GlobalizationSettings = new LocaleGlobalizationSettings(locale);

            // Scan all cells for error values
            List<string> errors = new List<string>();
            foreach (Worksheet sheet in wb.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate over used range only
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;
                for (int r = 0; r <= maxRow; r++)
                {
                    for (int c = 0; c <= maxCol; c++)
                    {
                        Cell cell = cells[r, c];
                        // Check if the cell contains an error value
                        if (cell.Type == CellValueType.IsError)
                        {
                            // Record the error string (e.g., "#DIV/0!")
                            errors.Add($"{sheet.Name}!{cell.Name}: {cell.StringValue}");
                        }
                    }
                }
            }

            // Store the collected information
            reports.Add(new WorkbookReportInfo
            {
                Path = Path.GetFullPath(filePath),
                Locale = locale,
                Errors = errors
            });
        }

        // Create a new workbook for the report using the provided create rule
        Workbook reportWb = new Workbook();
        Worksheet reportSheet = reportWb.Worksheets[0];
        reportSheet.Name = "LocalizationReport";

        // Write header row
        reportSheet.Cells["A1"].PutValue("Workbook Path");
        reportSheet.Cells["B1"].PutValue("Applied Locale");
        reportSheet.Cells["C1"].PutValue("Localization Errors");

        // Populate report rows
        int currentRow = 1; // zero‑based index; row 1 is the second row in the sheet
        foreach (WorkbookReportInfo info in reports)
        {
            reportSheet.Cells[currentRow, 0].PutValue(info.Path);
            reportSheet.Cells[currentRow, 1].PutValue(info.Locale);

            // Concatenate errors into a single string (separated by line breaks)
            string errorText = info.Errors.Count > 0 ? string.Join(Environment.NewLine, info.Errors) : "None";
            reportSheet.Cells[currentRow, 2].PutValue(errorText);

            currentRow++;
        }

        // Adjust column widths for readability
        reportSheet.AutoFitColumns();

        // Save the report workbook using the provided save rule
        reportWb.Save("LocalizationReport.xlsx");
    }
}