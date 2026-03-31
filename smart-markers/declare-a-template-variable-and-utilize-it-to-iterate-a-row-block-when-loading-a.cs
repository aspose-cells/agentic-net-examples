using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSX template that contains smart markers.
        // The template should have a variable marker like "&=$ReportTitle"
        // and a row block marker like "&=Employees" with column markers
        // "&=Employees.Name" and "&=Employees.Age".
        Workbook templateWorkbook = new Workbook("Template.xlsx");

        // Initialize WorkbookDesigner with the loaded workbook.
        WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

        // Declare a template variable named "ReportTitle" and assign a value.
        designer.SetDataSource("ReportTitle", "Employee Report");

        // Prepare a collection that will be used to iterate a row block.
        List<Employee> employees = new List<Employee>
        {
            new Employee { Name = "John Doe", Age = 30 },
            new Employee { Name = "Jane Smith", Age = 28 },
            new Employee { Name = "Bob Johnson", Age = 35 }
        };

        // Bind the collection to the smart marker "Employees".
        designer.SetDataSource("Employees", employees);

        // Process the smart markers. Setting LineByLine to false enables
        // range smart markers (recommended over the obsolete line‑by‑line mode).
        designer.LineByLine = false;
        designer.Process();

        // Save the populated workbook.
        templateWorkbook.Save("Result.xlsx");
    }

    // Simple data class used for the row block iteration.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}