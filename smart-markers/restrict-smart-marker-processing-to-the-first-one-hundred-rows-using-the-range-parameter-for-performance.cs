// Title: Restrict Aspose.Cells smart marker processing to the first 100 rows using a named Range in C#
// AI Prompts: Generate C# code that loads an Excel template, creates a DataTable with more than 100 rows, defines a named Aspose.Cells.Range for rows 2‑101, and calls WorkbookDesigner.Process on that range while preserving unknown markers. | Show how to configure WorkbookDesigner with a DataTable, create a range covering the first 100 data rows, and process only that range to improve smart marker performance. | Provide an example that saves the workbook after processing smart markers limited to a specific range and outputs a confirmation message.
// Common Searches: how to limit Aspose.Cells smart markers to first 100 rows in C# | using Aspose.Cells.Range to process only part of a worksheet for smart markers | performance tips for smart markers with large data tables in Aspose.Cells | C# example of partial smart marker processing with WorkbookDesigner | named range _CellsSmartMarkers requirement for Aspose.Cells smart markers
// Tags: process smart markers with Aspose.Cells range | limit smart marker rows C# | named range _CellsSmartMarkers Aspose.Cells | WorkbookDesigner partial range processing | smart marker performance optimization Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads or creates a workbook, builds a DataTable with 150 rows, defines a named range covering the first 100 rows, processes smart markers only within that range using WorkbookDesigner, preserves unrecognized markers, and saves the resulting file.
class RestrictSmartMarkerProcessing
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing template or create a new workbook if the file is missing
            Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a data source with more than 100 rows
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Salary", typeof(double));
            for (int i = 1; i <= 150; i++)
            {
                dt.Rows.Add($"Employee {i}", i * 1000);
            }

            // Set the data source for the designer
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource(dt);

            // Define a range that includes only the first 100 rows of data
            int startRow = 1;               // Row 2 in Excel (0‑based index)
            int startColumn = 0;            // Column A
            int totalRows = 100;            // First 100 rows
            int totalColumns = Math.Max(0, sheet.Cells.MaxDataColumn + 1); // All used columns

            // Use fully qualified Aspose.Cells.Range to avoid ambiguity with System.Range
            Aspose.Cells.Range range = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
            range.Name = "_CellsSmartMarkers"; // Required name for range smart markers

            // Process only the specified range; true = preserve unrecognized markers
            designer.Process(range, true);

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
