// Title: Create Excel Workbook with Nested Department‑Employee Data Using Range Smart Markers (C# Aspose.Cells)
// Description: This example shows how to build an Excel workbook, define a smart‑marker range (A2:C2), map parent (Department.Name) and child (Department.Employees.Name, Department.Employees.Age) fields, name the range _CellsSmartMarkers, bind a List<Department> to WorkbookDesigner, process the markers, and save the expanded sheet as an XLSX file.
// Keywords: Aspose.Cells | range smart markers | nested collections | C# Excel export | department employee hierarchy | WorkbookDesigner | smart marker range | Excel template | hierarchical data
// Common Searches: Aspose.Cells range smart markers nested collection | C# generate Excel from parent child list | How to use smart markers with List of objects | Aspose.Cells hierarchical data export example | Range smart markers .NET tutorial
// Developer Intent: Generate an Excel file that automatically expands rows for a parent‑child object hierarchy using range smart markers.
// Use Cases: Department‑wise employee directory where each department repeats for every employee. | Payroll or attendance sheets that group staff under their respective departments without manual row duplication. | Project task reports that list tasks and their sub‑tasks in a single template. | Export of any hierarchical business data (e.g., categories and products) to Excel with automatic row expansion.
// AI Prompts: Add a smart‑marker column that shows the total number of employees per department. | Insert a footer row after each department group that calculates the average age using range smart markers. | Extend the template to include an Employee Position column while preserving the nested hierarchy.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNestedSmartMarkersDemo
{
    // Sample data classes representing a nested hierarchy
    // This example shows how to build an Excel workbook, define a smart‑marker range (A2:C2), map parent (Department.Name) and child (Department.Employees.Name, Department.Employees.Age) fields, name the range _CellsSmartMarkers, bind a List<Department> to WorkbookDesigner, process the markers, and save the expanded sheet as an XLSX file.
    public class Department
    {
        public string Name { get; set; } = string.Empty;
        public List<Employee> Employees { get; set; } = new List<Employee>();
    }

    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Build the template using range smart markers
                // Header row
                sheet.Cells["A1"].PutValue("Department");
                sheet.Cells["B1"].PutValue("Employee Name");
                sheet.Cells["C1"].PutValue("Employee Age");

                // Data rows – smart markers that will be expanded automatically
                sheet.Cells["A2"].PutValue("&=Department.Name");                 // Parent collection field
                sheet.Cells["B2"].PutValue("&=Department.Employees.Name");       // Child collection field
                sheet.Cells["C2"].PutValue("&=Department.Employees.Age");        // Child collection field

                // Name the range that contains the smart markers (required for range smart markers)
                AsposeRange smartRange = sheet.Cells.CreateRange("A2:C2");
                smartRange.Name = "_CellsSmartMarkers";

                // 4. Prepare nested data source
                List<Department> departments = new List<Department>
                {
                    new Department
                    {
                        Name = "HR",
                        Employees = new List<Employee>
                        {
                            new Employee { Name = "John Doe", Age = 30 },
                            new Employee { Name = "Jane Smith", Age = 25 }
                        }
                    },
                    new Department
                    {
                        Name = "IT",
                        Employees = new List<Employee>
                        {
                            new Employee { Name = "Mike Brown", Age = 35 }
                        }
                    }
                };

                // 5. Create a WorkbookDesigner and assign the workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // 6. Set the data source. The name "Department" matches the smart‑marker prefix.
                designer.SetDataSource("Department", departments);

                // 7. Process the smart markers
                designer.Process();

                // 8. Save the result
                string outputPath = "NestedSmartMarkersOutput.xlsx";
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
