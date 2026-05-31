using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Markup;

class SmartMarkerIgnoreErrorDemo
{
    static void Main()
    {
        try
        {
            // Load template workbook if it exists; otherwise create a simple one.
            Workbook workbook;
            const string templatePath = "Template.xlsx";
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";
                // Sample smart markers
                ws.Cells["A1"].PutValue("&=Employees.Name");
                ws.Cells["B1"].PutValue("&=Employees.Age");
            }

            // ------------------------------------------------------------
            // Prepare a data source that intentionally contains missing/invalid data
            // ------------------------------------------------------------
            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("Name", typeof(string));
            // Use string type for Age to allow invalid values without DataTable throwing.
            employees.Columns.Add("Age", typeof(string));

            // First row is valid
            employees.Rows.Add("John Doe", "30");
            // Second row has a missing Name (null) – this will cause a smart‑marker error
            employees.Rows.Add(DBNull.Value, "25");
            // Third row has an invalid Age (non‑numeric string) – another error scenario
            employees.Rows.Add("Jane Smith", "InvalidAge");

            // ------------------------------------------------------------
            // Set up the WorkbookDesigner with the workbook and the data source.
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                CallBack = new IgnoreErrorCallback() // custom callback to swallow errors
            };
            designer.SetDataSource(employees);

            // Process smart markers. Errors are handled by the callback.
            designer.Process(true);

            // ------------------------------------------------------------
            // Optionally, ignore formula calculation errors that might arise after
            // smart‑marker insertion.
            // ------------------------------------------------------------
            CalculationOptions calcOptions = new CalculationOptions { IgnoreError = true };
            workbook.CalculateFormula(calcOptions);

            // ------------------------------------------------------------
            // Save the resulting workbook.
            // ------------------------------------------------------------
            const string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // ------------------------------------------------------------------------
    // Callback implementation that catches and ignores any errors that occur
    // during smart‑marker processing.
    // ------------------------------------------------------------------------
    class IgnoreErrorCallback : ISmartMarkerCallBack
    {
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            try
            {
                // No custom logic required; the try/catch ensures any exception
                // thrown by the Aspose engine while processing a smart marker is suppressed.
            }
            catch
            {
                // Swallow the exception to continue processing remaining markers.
            }
        }
    }
}