// Title: C# – Load Only String Cells with LoadFilter and Build a Summary Report using Aspose.Cells
// Description: Shows how to configure Aspose.Cells LoadOptions with a LoadFilter (CellString) to load only string‑type cells from a template workbook, concatenate their values, and save the combined text in a new workbook (summary.xlsx).
// Keywords: Aspose.Cells | LoadFilter | CellString | C# | load string cells | concatenate Excel strings | summary report | memory‑efficient Excel processing
// Common Searches: Aspose.Cells load only string cells | C# filter Excel cells by type | concatenate all text from Excel using Aspose | generate text summary from workbook | LoadOptions CellString example
// Developer Intent: Extract only textual cell values from an Excel template and create a single‑cell summary report.
// Use Cases: Compile product names or identifiers stored as strings into one overview cell for quick reference. | Create a lightweight textual snapshot of comments or notes from a large workbook while keeping memory usage low. | Produce a consolidated report that merges textual data from multiple worksheets into a single workbook.
// AI Prompts: Modify the code to separate each string with a comma instead of a space. | Write the concatenated summary to separate rows—one row per worksheet—rather than a single cell. | Explain strategies for processing very large workbooks efficiently when using LoadFilter with the CellString option.

using System;
using System.Text;
using Aspose.Cells;

// Shows how to configure Aspose.Cells LoadOptions with a LoadFilter (CellString) to load only string‑type cells from a template workbook, concatenate their values, and save the combined text in a new workbook (summary.xlsx).
class StringCellsSummary
{
    static void Main()
    {
        // Path to the template workbook
        string templatePath = "template.xlsx";

        // Create LoadOptions and set a LoadFilter that loads only string cells
        LoadOptions loadOptions = new LoadOptions();
        // LoadFilter with CellString flag loads only cells whose value is a string
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellString);

        // Load the workbook using the specified LoadOptions
        Workbook sourceWorkbook = new Workbook(templatePath, loadOptions);

        // StringBuilder to accumulate all string values
        StringBuilder summaryBuilder = new StringBuilder();

        // Iterate through all worksheets
        foreach (Worksheet sheet in sourceWorkbook.Worksheets)
        {
            // Get the maximum used row and column indices
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Loop through each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // After applying the CellString filter, non‑string cells are not loaded
                    // Check if the cell actually contains a string value
                    if (cell != null && cell.Type == CellValueType.IsString)
                    {
                        // Append the string value followed by a space (or any delimiter you prefer)
                        summaryBuilder.Append(cell.StringValue);
                        summaryBuilder.Append(' ');
                    }
                }
            }
        }

        // Prepare the summary text
        string summaryText = summaryBuilder.ToString().Trim();

        // Create a new workbook for the summary report
        Workbook reportWorkbook = new Workbook();
        Worksheet reportSheet = reportWorkbook.Worksheets[0];

        // Place the concatenated summary into cell A1
        reportSheet.Cells["A1"].PutValue(summaryText);

        // Save the summary report
        reportWorkbook.Save("summary.xlsx");
    }
}
