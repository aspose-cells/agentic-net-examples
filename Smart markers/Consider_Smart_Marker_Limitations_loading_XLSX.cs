using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SmartMarkerLoadingDemo
    {
        public static void Run()
        {
            // ---------- Load ----------
            // Disable Excel restriction checking while loading (e.g., long strings, oversized data)
            LoadOptions loadOptions = new LoadOptions
            {
                CheckExcelRestriction = false
            };

            // Load the XLSX template that contains smart markers
            Workbook workbook = new Workbook("template.xlsx", loadOptions);

            // Also turn off restriction checking for the workbook instance
            workbook.Settings.CheckExcelRestriction = false;

            // ---------- Process Smart Markers ----------
            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare a simple data source (DataTable) for demonstration
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Salary", typeof(double));
            dt.Rows.Add("John Doe", 12345.67);
            dt.Rows.Add("Jane Smith", 98765.43);

            // Assign the data source to the designer
            designer.SetDataSource(dt);

            // Process all smart markers; preserve unrecognized markers (true)
            designer.Process(true);

            // ---------- Save ----------
            // Configure save options to bypass Excel restrictions when writing the file
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                CheckExcelRestriction = false
            };

            // Save the processed workbook to an XLSX file
            workbook.Save("output.xlsx", saveOptions);
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}