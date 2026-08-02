using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsCalcModeExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Simulate database data using a DataTable
            // ------------------------------------------------------------
            DataTable dbTable = new DataTable("Employees");
            dbTable.Columns.Add("ID", typeof(int));
            dbTable.Columns.Add("Name", typeof(string));
            dbTable.Columns.Add("Salary", typeof(double));

            dbTable.Rows.Add(1, "John Doe", 50000);
            dbTable.Rows.Add(2, "Jane Smith", 62000);
            dbTable.Rows.Add(3, "Mike Johnson", 58000);

            // Create a data reader from the DataTable (acts like a DB IDataReader)
            using (IDataReader dataReader = dbTable.CreateDataReader())
            {
                // Define import options (show column names, insert rows, etc.)
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true,   // include column headers
                    InsertRows = true,         // insert rows if needed
                    ConvertNumericData = true, // convert numeric types automatically
                    DateFormat = "yyyy-MM-dd"
                };

                // Import data starting at cell A1 (row 0, column 0)
                cells.ImportData(dataReader, 0, 0, importOptions);
            }

            // ------------------------------------------------------------
            // Add some formulas that depend on the imported data
            // ------------------------------------------------------------
            // Example: total salary in column D (after the imported table)
            // Assuming headers occupy row 0, data starts at row 1, and Salary column is C (index 2)
            // Place formula in D2 (row 1, column 3) to sum the Salary column
            cells[1, 3].Formula = "=SUM(C2:C4)";

            // ------------------------------------------------------------
            // Disable automatic calculation and set manual mode
            // ------------------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            // Optional: ensure formulas are not calculated on open/save
            workbook.Settings.FormulaSettings.CalculateOnOpen = false;
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // ------------------------------------------------------------
            // Manually trigger calculation for consistency
            // ------------------------------------------------------------
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // Save the workbook (lifecycle rule)
            // ------------------------------------------------------------
            workbook.Save("ManualCalculationExample.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook created, data imported, and formulas calculated manually.");
        }
    }
}