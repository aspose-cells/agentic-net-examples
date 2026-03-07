using System;
using Aspose.Cells;
using System.Collections.Generic;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple POCO class used as a data source
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load the XLSX template that contains smart markers.
            //    The template can be created in Excel and should contain
            //    smart markers with parameters, for example:
            //      &=$Employees[?(@.Age>28)].Name   // returns names of employees older than 28
            //      &=$ReportDate                    // returns a scalar value
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook("template.xlsx");   // load existing XLSX file

            // -----------------------------------------------------------------
            // 2. Create a WorkbookDesigner and attach the loaded workbook.
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // -----------------------------------------------------------------
            // 3. Prepare data sources.
            //    a) A list of Employee objects – this will be used with a
            //       smart marker that contains a filter parameter.
            //    b) A scalar value for the report date.
            // -----------------------------------------------------------------
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "HR" },
                new Employee { Name = "Bob Johnson", Age = 35, Department = "IT" }
            };

            // Bind the list to the name "Employees" – this name is referenced in the smart markers.
            designer.SetDataSource("Employees", employees);

            // Bind a scalar value (report date) to the name "ReportDate".
            designer.SetDataSource("ReportDate", DateTime.Now.ToString("yyyy-MM-dd"));

            // -----------------------------------------------------------------
            // 4. Process the smart markers.
            //    The boolean parameter 'false' means that unrecognized smart
            //    markers will be removed (default behaviour).
            // -----------------------------------------------------------------
            designer.Process(false);

            // -----------------------------------------------------------------
            // 5. Save the populated workbook.
            // -----------------------------------------------------------------
            designer.Workbook.Save("output.xlsx");
        }
    }
}