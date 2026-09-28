// Title: Use a custom ICellsDataTable to bind hierarchical Department and Employee objects to WorkbookDesigner smart markers in Aspose.Cells for .NET
// AI Prompts: Generate a data table from a List<Department> and bind it to WorkbookDesigner using SetDataSource for smart marker processing. | Insert smart markers that reference &Dept.Name, &Dept.Employees.FirstName, and &Dept.Employees.LastName, then run WorkbookDesigner.Process to produce the final Excel file. | Demonstrate how CellsDataTableFactory automatically flattens nested collections so that hierarchical objects can be used with smart markers.
// Common Searches: asp.net core generate Excel with smart markers from hierarchical object list | how to bind nested collections to WorkbookDesigner in Aspose.Cells | using CellsDataTableFactory to flatten department employee data for smart markers | set custom data source for smart markers with complex objects Aspose.Cells .NET | example of smart markers using &Dept.Employees.FirstName in Aspose.Cells
// Tags: WorkbookDesigner custom data source for smart markers | smart markers hierarchical object binding Aspose.Cells | flatten nested collections with CellsDataTableFactory | generate Excel from Department and Employee classes | Aspose.Cells .NET hierarchical data example

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample defines Department and Employee classes, creates a list of departments, adds smart markers that reference the hierarchy, builds an ICellsDataTable via CellsDataTableFactory, binds it to WorkbookDesigner with SetDataSource, processes the markers, and saves the resulting workbook as an XLSX file.
public class Department
{
    public string Name { get; set; }
    public List<Employee> Employees { get; set; }

    public Department(string name, List<Employee> employees)
    {
        Name = name;
        Employees = employees;
    }
}

public class Employee
{
    public string FirstName { get; set; }
    public string LastName { get; set; }

    public Employee(string first, string last)
    {
        FirstName = first;
        LastName = last;
    }
}

public class CustomDataSourceDemo
{
    public static void Run()
    {
        try
        {
            // 1. Prepare complex hierarchical data
            var departments = new List<Department>
            {
                new Department(
                    "Sales",
                    new List<Employee>
                    {
                        new Employee("John", "Doe"),
                        new Employee("Jane", "Smith")
                    }),
                new Department(
                    "IT",
                    new List<Employee>
                    {
                        new Employee("Mike", "Brown"),
                        new Employee("Emily", "White")
                    })
            };

            // 2. Create a new workbook (or load a template if needed)
            Workbook workbook = new Workbook();

            // 3. Insert smart markers that reflect the hierarchy.
            //    Example:
            //    A1: &Dept.Name
            //    B1: &Dept.Employees.FirstName
            //    C1: &Dept.Employees.LastName
            Worksheet ws = workbook.Worksheets[0];
            ws.Cells["A1"].PutValue("&Dept.Name");
            ws.Cells["B1"].PutValue("&Dept.Employees.FirstName");
            ws.Cells["C1"].PutValue("&Dept.Employees.LastName");

            // 4. Create a WorkbookDesigner and assign the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // 5. Build an ICellsDataTable from the custom object list using the factory.
            //    The factory automatically flattens hierarchical collections for smart markers.
            ICellsDataTable dataTable = workbook.CellsDataTableFactory.GetInstance(departments);

            // 6. Bind the custom data source to a name used in the smart markers.
            designer.SetDataSource("Dept", dataTable);

            // 7. Process the smart markers
            designer.Process();

            // 8. Save the result
            string outputPath = "CustomHierarchicalDataOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook generated successfully: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error generating workbook: {ex.Message}");
        }
    }
}

class Program
{
    static void Main()
    {
        CustomDataSourceDemo.Run();
    }
}
