// Title: Insert a blank row between department groups with Aspose.Cells smart marker (group:normal,skip:1) in C#
// AI Prompts: Write C# code that uses Aspose.Cells WorkbookDesigner to group employee rows by Department and inserts a blank row after each group using the (group:normal,skip:1) smart marker option. | Show how to assign a List<Employee> as the data source for smart markers and apply grouping with a blank separator row in Aspose.Cells. | Generate an example that creates an Excel workbook, places smart marker expressions for grouping, processes them, and saves the file with empty rows between each department group.
// Common Searches: asp.net how to add empty rows between grouped smart marker sections in Aspose.Cells | c# Aspose.Cells smart marker group normal skip 1 blank line example | using WorkbookDesigner to group rows and insert spacing rows in Excel export | smart marker syntax for inserting a blank row after each department group | Aspose.Cells grouping data with blank rows between groups C#
// Tags: smart marker grouping blank rows | WorkbookDesigner set list data source | Aspose.Cells group normal skip option | C# Excel export grouped data | insert empty row between groups Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for WorkbookDesigner namespace

namespace AsposeCellsSmartMarkerGrouping
{
    // Simple data class used as the data source for smart markers
    // The example creates a workbook, adds smart marker cells with '&=Employees.Group:normal,skip:1' to group rows by Department, inserts a blank row after each group, sets a List<Employee> as the data source, processes the markers via WorkbookDesigner, and saves the result as GroupedData.xlsx.
    public class Employee
    {
        public string Department { get; set; }
        public string Name { get; set; }

        public Employee(string department, string name)
        {
            Department = department;
            Name = name;
        }
    }

    public class GroupDataWithSmartMarker
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // 2. Set up header cells
            worksheet.Cells["A1"].PutValue("Department");
            worksheet.Cells["B1"].PutValue("Employee Name");

            // 3. Insert smart marker expressions.
            //    The "(group:normal,skip:1)" part tells Aspose.Cells to insert a blank row
            //    after each group (i.e., after each distinct Department value).
            worksheet.Cells["A2"].PutValue("&=Employees.Group:normal,skip:1");
            worksheet.Cells["B2"].PutValue("&=Employees.Name");

            // 4. Prepare the data source – a list of employees with department information.
            List<Employee> employees = new List<Employee>
            {
                new Employee("Sales", "John Doe"),
                new Employee("Sales", "Jane Smith"),
                new Employee("Marketing", "Alice Johnson"),
                new Employee("Marketing", "Bob Brown"),
                new Employee("IT", "Charlie Davis")
            };

            // 5. Create a WorkbookDesigner, assign the workbook, and set the data source.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", employees);

            // 6. Process all smart markers in the workbook.
            //    This will expand the smart marker range, group rows by Department,
            //    and insert a blank row between each group as specified.
            designer.Process();

            // 7. Save the resulting workbook.
            workbook.Save("GroupedData.xlsx");
        }
    }
}
