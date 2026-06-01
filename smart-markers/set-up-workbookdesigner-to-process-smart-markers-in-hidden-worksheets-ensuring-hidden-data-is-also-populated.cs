using System;
using System.Collections.Generic;
using Aspose.Cells;

class ProcessHiddenSmartMarkers
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Add a hidden worksheet that will contain the smart markers
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenData");
        hiddenSheet.IsVisible = false; // hide the worksheet

        // Insert smart markers into the hidden sheet
        hiddenSheet.Cells["A1"].PutValue("&=Employees.Name");
        hiddenSheet.Cells["B1"].PutValue("&=Employees.Age");

        // Optional: add a visible sheet to display the populated data
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "Report";
        visibleSheet.Cells["A1"].PutValue("Name");
        visibleSheet.Cells["B1"].PutValue("Age");
        // Use the same smart markers on the visible sheet (they will be filled from the same data source)
        visibleSheet.Cells["A2"].PutValue("&=Employees.Name");
        visibleSheet.Cells["B2"].PutValue("&=Employees.Age");

        // Prepare a data source that matches the smart marker table name "Employees"
        List<Employee> employees = new List<Employee>
        {
            new Employee { Name = "John Doe", Age = 30 },
            new Employee { Name = "Jane Smith", Age = 28 }
        };

        // Set up WorkbookDesigner (lifecycle: create)
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Ensure that any formulas referencing the hidden sheet are updated after processing
        designer.UpdateReference = true;

        // Bind the data source to the smart marker name
        designer.SetDataSource("Employees", employees);

        // Process all smart markers in the workbook, including those in hidden worksheets
        designer.Process();

        // Save the processed workbook (lifecycle: save)
        workbook.Save("ProcessedHiddenSmartMarkers.xlsx");
    }

    // Simple POCO class representing the data structure for the smart markers
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}