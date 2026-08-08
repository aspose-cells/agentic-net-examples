// Title: C# Unit Test for Aspose.Cells Smart Marker Replacement – Verify Cell Values
// Description: Demonstrates how to create a C# unit test that builds a workbook with smart markers, binds a DataTable, processes the markers using WorkbookDesigner, asserts that cells A1‑B2 contain the expected employee names and ages, checks that the marker syntax is removed, and optionally saves the file to a temporary location.
// Keywords: Aspose.Cells | smart markers | C# unit test | WorkbookDesigner | verify cell values | .NET testing | data binding | Excel automation
// Common Searches: Aspose.Cells unit test smart markers C# | how to assert smart marker results in .NET | verify workbookdesigner output cells | check smart marker removal after processing | C# test for Excel smart marker replacement
// Developer Intent: Create an automated test that confirms smart marker processing populates the worksheet with the correct data and eliminates the marker placeholders.
// Use Cases: Validate that a DataTable bound to the "Employees" smart marker fills A1:B2 with the correct names and ages. | Ensure no residual "&=" syntax remains after WorkbookDesigner processes the workbook. | Confirm the generated workbook can be saved and opened without errors in a CI pipeline.
// AI Prompts: Generate an MSTest method that uses Aspose.Cells WorkbookDesigner to process smart markers and asserts the resulting cell values and marker removal. | Provide a NUnit test example that binds a DataTable to smart markers, runs the designer, and validates the output cells and absence of "&=" syntax. | Create an xUnit test that processes smart markers in a workbook, checks cells A1‑B2 against expected data, and confirms the workbook can be saved to a temporary file.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsSmartMarkerTests
{
    // Implements a simple execution that verifies smart marker replacement results.
    // Demonstrates how to create a C# unit test that builds a workbook with smart markers, binds a DataTable, processes the markers using WorkbookDesigner, asserts that cells A1‑B2 contain the expected employee names and ages, checks that the marker syntax is removed, and optionally saves the file to a temporary location.
    public class SmartMarkerReplacementDemo
    {
        public static void Main()
        {
            try
            {
                RunDemo();
                Console.WriteLine("Smart marker replacement demo completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void RunDemo()
        {
            // ---------- Create a workbook with smart markers ----------
            Workbook workbook = new Workbook();                         // create workbook
            Worksheet sheet = workbook.Worksheets[0];                  // get first worksheet

            // Place smart markers in cells. The syntax "&=Table.Column" tells Aspose.Cells to replace with data.
            sheet.Cells["A1"].PutValue("&=Employees.Name");
            sheet.Cells["B1"].PutValue("&=Employees.Age");

            // ---------- Prepare data source ----------
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // ---------- Set up WorkbookDesigner ----------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // Bind the DataTable to the smart marker name "Employees"
            designer.SetDataSource("Employees", dt);
            // Process the smart markers (populate the worksheet)
            designer.Process();

            // ---------- Verify replacement results ----------
            // After processing, the first row (A1,B1) should contain the first data row,
            // and the second row (A2,B2) should contain the second data row.
            if (sheet.Cells["A1"].StringValue != "John Doe" ||
                sheet.Cells["B1"].IntValue != 30 ||
                sheet.Cells["A2"].StringValue != "Jane Smith" ||
                sheet.Cells["B2"].IntValue != 28)
            {
                throw new InvalidOperationException("Smart marker replacement did not produce expected results.");
            }

            // ---------- Optional: Ensure no leftover smart markers ----------
            if (sheet.Cells["A1"].StringValue.Contains("&=") ||
                sheet.Cells["B1"].StringValue.Contains("&="))
            {
                throw new InvalidOperationException("Smart markers were not fully removed after processing.");
            }

            // ---------- Save workbook (demonstrates usage of save rule) ----------
            // The file is saved to a temporary location; in real unit tests this may be omitted.
            string tempPath = Path.GetTempFileName().Replace(".tmp", ".xlsx");
            workbook.Save(tempPath);
            Console.WriteLine($"Workbook saved to: {tempPath}");
        }
    }
}
