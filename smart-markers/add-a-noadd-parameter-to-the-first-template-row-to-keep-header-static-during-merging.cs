// Title: Keep Header Row Static Using NoAdd in Aspose.Cells Smart Markers – C# Example
// Description: Shows how to use the `noadd` directive (via ImportTableOptions.ShiftFirstRowDown) on the first template row so the header remains unchanged when importing a DataTable and merging cells with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | noadd | smart markers | C# | ImportTableOptions | ShiftFirstRowDown | static header | merge cells | DataTable import | Excel report template
// Common Searches: Aspose.Cells noadd header row | C# keep header static smart markers | ImportTableOptions ShiftFirstRowDown example | merge rows without moving header Aspose.Cells | smart marker noadd parameter usage
// Developer Intent: Apply the `noadd` attribute to the first template row so the header is not duplicated or shifted during data import and subsequent merge operations.
// Use Cases: Import a DataTable that already contains column titles and add data rows without altering the header. | Generate Excel reports where column headings stay fixed while data rows are dynamically added and merged. | Create templates that merge cells in data rows but keep the first row untouched by using the noadd setting.
// AI Prompts: Provide C# code that uses Aspose.Cells smart markers with a noadd directive on the first template row to keep the header static during a merge. | Explain the difference between the noadd attribute and ShiftFirstRowDown when importing a DataTable with a header in Aspose.Cells for .NET. | Show how to merge cells below a static header after importing data with the noadd parameter in Aspose.Cells.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to use the `noadd` directive (via ImportTableOptions.ShiftFirstRowDown) on the first template row so the header remains unchanged when importing a DataTable and merging cells with Aspose.Cells for .NET.
    public class NoAddHeaderDuringMergeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Prepare a DataTable where the first row is the header
                DataTable table = new DataTable();
                table.Columns.Add("Header1");
                table.Columns.Add("Header2");
                table.Rows.Add("Header1", "Header2"); // header row (template row)
                table.Rows.Add("Data1", "Data2");
                table.Rows.Add("Data3", "Data4");

                // Configure ImportTableOptions:
                // ShiftFirstRowDown = true ensures the first template row (header) stays static
                // and new rows are inserted below it during the import/merge operation.
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    ShiftFirstRowDown = true,
                    IsFieldNameShown = false // we already have the header in the DataTable
                };

                // Import the DataTable starting at cell A1
                cells.ImportData(table, 0, 0, importOptions);

                // Example of merging cells below the header to demonstrate that the header remains unchanged
                // Merge cells B2:C2 (second data row) – the header in row 1 stays intact.
                cells.Merge(1, 1, 1, 2);
                cells[1, 1].PutValue("Merged Data");

                // Save the workbook
                workbook.Save("NoAddHeaderDuringMergeDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
