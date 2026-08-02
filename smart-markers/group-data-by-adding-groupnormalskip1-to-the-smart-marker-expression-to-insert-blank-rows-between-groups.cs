using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a new workbook that will serve as the template.
        // ------------------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add column headers.
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Score");

        // Insert smart markers with grouping syntax.
        // (group:normal,skip:1) groups rows by the data source and inserts a blank row between each group.
        sheet.Cells["A2"].PutValue("&=Employees.Name(group:normal,skip:1)");
        sheet.Cells["B2"].PutValue("&=Employees.Score(group:normal,skip:1)");

        // Define the range that contains the smart markers.
        // The range must be named "_CellsSmartMarkers" for the designer to recognize it.
        sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

        // ------------------------------------------------------------
        // 2. Prepare the data source.
        // ------------------------------------------------------------
        List<Employee> employees = new List<Employee>
        {
            new Employee { Name = "Alice",   Score = 85 },
            new Employee { Name = "Bob",     Score = 90 },
            new Employee { Name = "Charlie", Score = 78 },
            new Employee { Name = "David",   Score = 92 }
        };

        // ------------------------------------------------------------
        // 3. Set up the WorkbookDesigner, bind the data source, and process.
        // ------------------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Employees", employees);
        designer.Process(); // processes all smart markers in the workbook

        // ------------------------------------------------------------
        // 4. Save the resulting workbook.
        // ------------------------------------------------------------
        workbook.Save("GroupedSmartMarkers.xlsx");
    }

    // Simple POCO class representing each row of data.
    public class Employee
    {
        public string Name { get; set; }
        public int Score { get; set; }
    }
}