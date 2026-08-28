// Title: Create grouped Excel rows from Department and Employee objects by disabling WorkbookDesigner.LineByLine in Aspose.Cells for .NET
// AI Prompts: Generate an Excel workbook in C# that binds a List<Department> containing nested Employee lists and outputs grouped rows by setting WorkbookDesigner.LineByLine = false. | Define a named smart‑marker range called _CellsSmartMarkers and process the template with WorkbookDesigner to merge parent‑child data into a single record set.
// Common Searches: Aspose.Cells how to disable line‑by‑line processing for nested collections | C# smart markers grouped records with parent child objects | WorkbookDesigner set LineByLine false example with named range | Export Department and Employee data to Excel using Aspose.Cells smart markers | named smart marker range _CellsSmartMarkers usage in Aspose.Cells
// Tags: WorkbookDesigner line-by-line false | smart markers nested collection | named smart marker range _CellsSmartMarkers | grouped records Aspose.Cells | C# export parent child objects to Excel

using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsLineByLineDemo
{
    // Sample child class representing nested objects
    // Demonstrates using WorkbookDesigner with LineByLine set to false and a named smart‑marker range to merge Department objects and their Employee collections into grouped records in an Excel workbook.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Sample parent class containing a collection of nested objects
    public class Department
    {
        public string DeptName { get; set; } = string.Empty;
        public List<Employee> Employees { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Insert smart markers for a parent object (Department) and its nested collection (Employees)
                //    The markers are placed inside a range that will be named "_CellsSmartMarkers"
                sheet.Cells["A1"].PutValue("&Department.DeptName");          // Parent property
                sheet.Cells["A2"].PutValue("&Department.Employees.Name");   // Nested collection property
                sheet.Cells["B2"].PutValue("&Department.Employees.Age");    // Nested collection property

                // Define the range that contains the smart markers and give it the required name
                // This is needed when LineByLine is set to false
                AsposeRange smartMarkerRange = sheet.Cells.CreateRange("A1:B2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // 3. Prepare sample data with nested objects
                List<Department> departments = new List<Department>
                {
                    new Department
                    {
                        DeptName = "Sales",
                        Employees = new List<Employee>
                        {
                            new Employee { Name = "John Doe", Age = 30 },
                            new Employee { Name = "Jane Smith", Age = 28 }
                        }
                    },
                    new Department
                    {
                        DeptName = "HR",
                        Employees = new List<Employee>
                        {
                            new Employee { Name = "Bob Johnson", Age = 35 }
                        }
                    }
                };

                // 4. Initialize WorkbookDesigner, assign the workbook, and set LineByLine to false
                //    This tells the designer to treat the nested collection as a grouped record set
                WorkbookDesigner designer = new WorkbookDesigner(workbook)
                {
                    LineByLine = false // Obsolete but still functional; using range smart markers instead of line‑by‑line
                };

                // 5. Bind the data source to the smart marker name used in the template
                designer.SetDataSource("Department", departments);

                // 6. Process the smart markers
                designer.Process();

                // 7. Save the resulting workbook
                string outputPath = "GroupedNestedOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
