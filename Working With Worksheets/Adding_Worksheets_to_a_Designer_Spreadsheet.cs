using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDesignerExample
{
    // Simple POCO class to be used as a data source
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Employee(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    public class AddWorksheetsToDesigner
    {
        public static void Run()
        {
            // 1. Create a new workbook (uses Workbook() constructor)
            Workbook workbook = new Workbook();

            // 2. Add a new worksheet to the workbook collection with a custom name
            //    (uses WorksheetCollection.Add(string) method)
            Worksheet dataSheet = workbook.Worksheets.Add("DataSheet");

            // 3. Insert smart markers into the new worksheet.
            //    These markers will be replaced by the data source during processing.
            dataSheet.Cells["A1"].PutValue("&Employee.Name");
            dataSheet.Cells["B1"].PutValue("&Employee.Age");

            // 4. Prepare a sample data source (a list of Employee objects)
            List<Employee> employees = new List<Employee>
            {
                new Employee("John Doe", 30),
                new Employee("Jane Smith", 28),
                new Employee("Bob Johnson", 45)
            };

            // 5. Create a WorkbookDesigner and assign the workbook to it
            //    (uses WorkbookDesigner() constructor and WorkbookDesigner.Workbook property)
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // 6. Bind the data source to the smart marker name "Employee"
            //    (uses WorkbookDesigner.SetDataSource(string, object) method)
            designer.SetDataSource("Employee", employees);

            // 7. Process the smart markers so that the data is populated into the worksheet
            //    (uses WorkbookDesigner.Process() method)
            designer.Process();

            // 8. Save the resulting workbook to disk
            //    (uses Workbook.Save(string) method)
            string outputPath = "DesignerWithAddedWorksheet.xlsx";
            designer.Workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main(string[] args)
        {
            AddWorksheetsToDesigner.Run();
        }
    }
}