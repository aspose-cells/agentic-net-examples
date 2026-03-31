using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerProcessingDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook with specific LoadOptions.
            // CheckDataValid = false ensures that any data validation errors in the template
            // do not prevent smart marker processing.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CheckDataValid = false,
                // Keep formula parsing enabled so that smart marker formulas are recognized.
                ParsingFormulaOnOpen = true
            };

            // Replace "template.xlsx" with the path to your XLSX template containing smart markers.
            Workbook workbook = new Workbook("template.xlsx", loadOptions);

            // ------------------------------------------------------------
            // Prepare a data source for the smart markers.
            // In this example we use a DataTable, but any supported source can be used.
            // ------------------------------------------------------------
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Department", typeof(string));

            dt.Rows.Add("John Doe", 30, "Sales");
            dt.Rows.Add("Jane Smith", 28, "Marketing");
            dt.Rows.Add("Bob Johnson", 45, "HR");

            // ------------------------------------------------------------
            // Set up the WorkbookDesigner to process smart markers.
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // When LineByLine is false, the template must contain a named range "_CellsSmartMarkers".
                // This allows range‑based smart marker processing (the recommended approach).
                LineByLine = false
            };

            // Assign the data source to a name that matches the smart marker prefix in the template.
            designer.SetDataSource("Employees", dt);

            // ------------------------------------------------------------
            // Process the smart markers.
            // Since LineByLine = false, the designer will look for the named range.
            // If the range is not defined, you can alternatively call designer.Process()
            // which works with the older line‑by‑line mode.
            // ------------------------------------------------------------
            designer.Process();

            // ------------------------------------------------------------
            // Save the processed workbook.
            // ------------------------------------------------------------
            workbook.Save("ProcessedOutput.xlsx");
        }
    }
}