// Title: C# – Load only string cells with Aspose.Cells LoadFilter and generate a concatenated summary workbook
// Description: Demonstrates how to use Aspose.Cells LoadFilter with LoadDataFilterOptions.CellString to load only text cells from a template Excel file, concatenate their values, and write the combined string to a new workbook as a summary report. The example improves performance by skipping non‑string data and creates a single‑cell overview of all textual content.
// Keywords: Aspose.Cells | LoadFilter | CellString | C# | .NET | load string cells | concatenate cell values | summary report workbook | Excel performance | filter Excel data | string extraction
// Common Searches: Aspose.Cells load only string cells | C# concatenate all text from Excel workbook | How to use LoadFilter CellString in Aspose.Cells | Create summary report from Excel strings .NET | Improve Excel read performance with LoadFilter
// Developer Intent: Load only string‑type cells from a template workbook and combine them into a single summary report.
// Use Cases: Produce a quick‑read overview of product descriptions spread across multiple sheets. | Reduce memory usage and processing time when extracting text from large spreadsheets. | Generate a printable one‑cell summary that can be exported or emailed.
// AI Prompts: Show how to change the separator from a space to a comma when concatenating string values. | Provide code to split the concatenated text into multiple rows if it exceeds a given length. | Explain how to combine LoadDataFilterOptions.CellString with other filter options to load mixed data types selectively.

using System;
using System.Text;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells LoadFilter with LoadDataFilterOptions.CellString to load only text cells from a template Excel file, concatenate their values, and write the combined string to a new workbook as a summary report. The example improves performance by skipping non‑string data and creates a single‑cell overview of all textual content.
class Program
{
    static void Main()
    {
        // Path to the template workbook containing data
        string templatePath = "Template.xlsx";

        // Path where the summary report will be saved
        string summaryPath = "SummaryReport.xlsx";

        // Create a LoadFilter that loads only cells with string values
        LoadFilter loadFilter = new LoadFilter(LoadDataFilterOptions.CellString);
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFilter = loadFilter
        };

        // Load the workbook using the filter options
        Workbook sourceWorkbook = new Workbook(templatePath, loadOptions);

        // StringBuilder to accumulate all string values
        StringBuilder concatenatedStrings = new StringBuilder();

        // Iterate through each worksheet and its used range
        foreach (Worksheet sheet in sourceWorkbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    // Only process cells that are of string type
                    if (cell.Type == CellValueType.IsString)
                    {
                        concatenatedStrings.Append(cell.StringValue);
                        concatenatedStrings.Append(' '); // separator between values
                    }
                }
            }
        }

        // Create a new workbook to hold the summary report
        Workbook summaryWorkbook = new Workbook();
        Worksheet summarySheet = summaryWorkbook.Worksheets[0];

        // Write the concatenated string into cell A1 of the summary sheet
        summarySheet.Cells["A1"].PutValue(concatenatedStrings.ToString().Trim());

        // Save the summary workbook
        summaryWorkbook.Save(summaryPath);
    }
}
