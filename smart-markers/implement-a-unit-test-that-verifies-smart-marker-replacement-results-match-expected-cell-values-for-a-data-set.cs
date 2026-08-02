// Title: C# Unit Test for Aspose.Cells Smart Marker Replacement with Employee List
// Description: Creates a workbook, places smart markers "&=Employees.Name" and "&=Employees.Age" in row 2, binds a List<Employee> containing two records, processes the markers with WorkbookDesigner, and asserts that cells A2, B2, A3 and B3 contain the expected names and ages.
// Keywords: Aspose.Cells | Smart Markers | C# unit test | WorkbookDesigner | data binding | employee list | cell value verification | .NET testing | NUnit | MSTest | xUnit
// Common Searches: Aspose.Cells smart marker unit test C# | verify smart marker output with WorkbookDesigner | unit testing Aspose.Cells row expansion | C# test for smart marker data binding | how to assert smart marker results in .NET
// Developer Intent: Write an automated test that confirms smart marker processing expands rows correctly and populates cells with the expected employee names and ages.
// Use Cases: Automated regression testing of smart‑marker templates before release. | Ensuring that a collection bound to smart markers generates the correct number of rows and values. | Validating data‑binding logic in report generation pipelines that use Aspose.Cells.
// AI Prompts: Generate an NUnit test method that creates a workbook, inserts smart markers, binds a List<Employee>, processes the markers, and asserts the cell values for each employee. | Provide a MSTest example that verifies Aspose.Cells smart marker replacement for a two‑item employee list. | Write a xUnit test that checks row expansion and cell contents after WorkbookDesigner processes smart markers.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerTests
{
    // Simple data class used as a data source for smart markers
    // Creates a workbook, places smart markers "&=Employees.Name" and "&=Employees.Age" in row 2, binds a List<Employee> containing two records, processes the markers with WorkbookDesigner, and asserts that cells A2, B2, A3 and B3 contain the expected names and ages.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class SmartMarkerReplacement
    {
        public static void Main()
        {
            try
            {
                RunSmartMarkerTest();
                Console.WriteLine("Smart marker replacement succeeded.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during smart marker processing: {ex.Message}");
            }
        }

        private static void RunSmartMarkerTest()
        {
            // ---------- Create ----------
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Place smart markers in the template.
            // Row 2 will be the data row that will be repeated for each Employee.
            // "&=Employees.Name" and "&=Employees.Age" are the smart markers.
            sheet.Cells["A2"].PutValue("&=Employees.Name");
            sheet.Cells["B2"].PutValue("&=Employees.Age");

            // ---------- Prepare Data ----------
            // Create a list of employees that will be bound to the smart markers.
            var employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 30 },
                new Employee { Name = "Jane Smith", Age = 28 }
            };

            // ---------- Process Smart Markers ----------
            // Set up the WorkbookDesigner, bind the data source and process the markers.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // Bind the list to the name used in the smart markers ("Employees")
            designer.SetDataSource("Employees", employees);
            // Process the smart markers (default is line‑by‑line processing)
            designer.Process();

            // ---------- Verify Results ----------
            // After processing, the smart marker row should be expanded to two rows:
            // Row 2 -> first employee, Row 3 -> second employee.
            if (sheet.Cells["A2"].StringValue != "John Doe")
                throw new InvalidOperationException("First employee name mismatch.");
            if (sheet.Cells["B2"].IntValue != 30)
                throw new InvalidOperationException("First employee age mismatch.");

            if (sheet.Cells["A3"].StringValue != "Jane Smith")
                throw new InvalidOperationException("Second employee name mismatch.");
            if (sheet.Cells["B3"].IntValue != 28)
                throw new InvalidOperationException("Second employee age mismatch.");

            // ---------- Save (optional) ----------
            // The workbook can be saved to verify manually if needed.
            // workbook.Save("SmartMarkerResult.xlsx");
        }
    }
}
