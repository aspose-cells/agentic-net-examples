// Title: C# unit test that verifies Aspose.Cells smart marker replacement using a DataTable data source
// AI Prompts: Create an MSTest method that builds a workbook, inserts '&=Employees.Name' and '&=Employees.Age' smart markers, assigns a DataTable named Employees as the data source via WorkbookDesigner, processes the markers, and asserts that cells A1, B1, A2, and B2 contain the expected values. | Generate a NUnit test that loads a workbook with smart markers, calls WorkbookDesigner.Process, and uses Assert.AreEqual to compare the string values of the resulting cells with the original DataTable rows.
// Common Searches: how to unit test Aspose.Cells smart markers with a DataTable in C# | Aspose.Cells verify smart marker output using MSTest | C# example of asserting cell values after WorkbookDesigner processing | unit testing smart marker replacement results in Aspose.Cells .NET | write a test for WorkbookDesigner SetDataSource and smart marker expansion
// Tags: Aspose.Cells WorkbookDesigner SetDataSource unit test | smart marker replacement verification C# | assert cell values after smart marker processing | NUnit test for Aspose.Cells smart markers | DataTable as smart marker data source Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // The example creates a workbook, places smart markers for an Employees DataTable, processes them with WorkbookDesigner, and throws exceptions if the resulting cell values (A1, B1, A2, B2) do not match the expected names and ages, serving as a basis for a unit test.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Insert smart markers for a collection named "Employees"
                // These markers will be replaced with the data rows during processing
                cells["A1"].PutValue("&=Employees.Name");
                cells["B1"].PutValue("&=Employees.Age");

                // Prepare the data source: a DataTable with two columns and two rows
                DataTable employeeTable = new DataTable("Employees");
                employeeTable.Columns.Add("Name", typeof(string));
                employeeTable.Columns.Add("Age", typeof(int));

                employeeTable.Rows.Add("John", 30);
                employeeTable.Rows.Add("Jane", 25);

                // Set the data source for the smart marker processor
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                // Use the overload that accepts a name and an object (DataTable)
                designer.SetDataSource("Employees", employeeTable);
                designer.Process();

                // Verify that the smart markers were replaced with the expected values
                // Row 0 (A1, B1) should contain the first employee's data
                if (cells["A1"].StringValue != "John")
                    throw new Exception($"Expected 'John' in A1 but found '{cells["A1"].StringValue}'.");

                if (cells["B1"].StringValue != "30")
                    throw new Exception($"Expected '30' in B1 but found '{cells["B1"].StringValue}'.");

                // Row 1 (A2, B2) should contain the second employee's data
                if (cells["A2"].StringValue != "Jane")
                    throw new Exception($"Expected 'Jane' in A2 but found '{cells["A2"].StringValue}'.");

                if (cells["B2"].StringValue != "25")
                    throw new Exception($"Expected '25' in B2 but found '{cells["B2"].StringValue}'.");

                Console.WriteLine("Smart marker replacement verified successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
