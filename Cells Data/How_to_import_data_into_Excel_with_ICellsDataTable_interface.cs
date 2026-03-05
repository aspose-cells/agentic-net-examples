using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            ImportDataWithICellsDataTable.Run();
        }
    }

    public static class ImportDataWithICellsDataTable
    {
        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Prepare sample data as a list of rows (each row is an object array)
            ArrayList dataLists = new ArrayList();

            // Header row
            dataLists.Add(new object[] { "Name", "Age", "Gender" });

            // Data rows
            dataLists.Add(new object[] { "Alice", 30, "Female" });
            dataLists.Add(new object[] { "Bob", 25, "Male" });
            dataLists.Add(new object[] { "Charlie", 35, "Male" });

            // Convert the list to an ICellsDataTable instance
            ICellsDataTable dataTable = workbook.CellsDataTableFactory.GetInstance(dataLists, true);

            // Import the data table into the first worksheet starting at cell A1 (row 0, column 0)
            workbook.Worksheets[0].Cells.ImportData(dataTable, 0, 0, new ImportTableOptions());

            // Verify import by reading a cell value
            Console.WriteLine("Value at B2 (should be 30): " + workbook.Worksheets[0].Cells["B2"].StringValue);

            // Save the workbook in XLSX format
            workbook.Save("ICellsDataTableImportDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}