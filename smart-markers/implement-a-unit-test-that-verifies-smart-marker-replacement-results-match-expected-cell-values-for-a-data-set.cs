// Title: Create a C# unit test with Aspose.Cells to validate smart marker replacement from a DataTable
// AI Prompts: Write an MSTest method that builds a workbook, inserts smart markers, binds a DataTable as a data source, processes the markers with WorkbookDesigner, and asserts the expected values in each generated cell. | Generate a NUnit test case that creates a workbook, adds smart markers, sets the DataTable source, runs Process(), and uses Assert.AreEqual to verify cell contents for all rows. | Provide an xUnit test that performs the same smart‑marker workflow, validates each cell value, and optionally saves the workbook to a MemoryStream for further inspection.
// Common Searches: how to write a unit test for Aspose.Cells smart markers in C# | Aspose.Cells WorkbookDesigner unit test example with DataTable | C# test verifying smart marker values after processing | unit testing Excel smart markers using MSTest NUnit or xUnit | assert cell values in Aspose.Cells after smart marker replacement
// Tags: Aspose.Cells smart marker unit test | C# WorkbookDesigner data source validation | verify smart marker replacement in Excel | assert cell values after smart marker processing | DataTable to smart marker conversion test

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Demonstrates how to write a C# unit test that creates a workbook, places smart markers, binds a DataTable as the data source, processes the markers with WorkbookDesigner, and asserts that each resulting cell contains the expected value, optionally saving the workbook to a memory stream.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a workbook with smart markers ----------
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Insert smart markers in the first row.
                // "&=Data.Name" and "&=Data.Age" will be replaced by the data source values.
                cells["A1"].PutValue("&=Data.Name");
                cells["B1"].PutValue("&=Data.Age");

                // ---------- Prepare the data source ----------
                DataTable dt = new DataTable("Data");
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Age", typeof(int));

                dt.Rows.Add("Alice", 30);
                dt.Rows.Add("Bob", 25);
                dt.Rows.Add("Charlie", 35);

                // ---------- Process smart markers ----------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Data", dt);
                designer.Process();

                // ---------- Verify the results ----------
                ValidateCell(cells["A1"], "Alice");
                ValidateCell(cells["B1"], 30);
                ValidateCell(cells["A2"], "Bob");
                ValidateCell(cells["B2"], 25);
                ValidateCell(cells["A3"], "Charlie");
                ValidateCell(cells["B3"], 35);

                // ---------- Optional: Save the workbook to a memory stream ----------
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    // The stream now contains the generated Excel file.
                    // No further action needed for this demo.
                }

                Console.WriteLine("Smart marker replacement test passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Simple validation helper that throws if the cell value does not match the expected value.
        private static void ValidateCell(Cell cell, object expected)
        {
            object actual = cell.Value;
            if (!object.Equals(actual, expected))
            {
                throw new InvalidOperationException(
                    $"Validation failed for cell {cell.Name}. Expected: {expected}, Actual: {actual}");
            }
        }
    }
}
