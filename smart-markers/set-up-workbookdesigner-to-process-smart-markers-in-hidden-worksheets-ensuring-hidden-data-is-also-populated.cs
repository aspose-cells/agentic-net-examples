using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsHiddenSmartMarkersDemo
{
    // Sample data class used as a data source
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (template) and add sheets
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet visibleSheet = workbook.Worksheets[0];        // first sheet (visible by default)
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenData"); // add a second sheet

            // Hide the second worksheet
            hiddenSheet.IsVisible = false; // keep it hidden after processing

            // -------------------------------------------------
            // 2. Insert smart markers into both worksheets
            // -------------------------------------------------
            // Visible sheet smart markers (line‑by‑line processing disabled)
            visibleSheet.Cells["A1"].PutValue("&=Employees.Name");
            visibleSheet.Cells["B1"].PutValue("&=Employees.Age");
            visibleSheet.Cells["C1"].PutValue("&=Employees.Department");
            // Mark the range that contains smart markers (required when LineByLine = false)
            visibleSheet.Cells.CreateRange("A1:C1").Name = "_CellsSmartMarkers";

            // Hidden sheet smart markers
            hiddenSheet.Cells["A1"].PutValue("&=Employees.Name");
            hiddenSheet.Cells["B1"].PutValue("&=Employees.Age");
            hiddenSheet.Cells["C1"].PutValue("&=Employees.Department");
            hiddenSheet.Cells.CreateRange("A1:C1").Name = "_CellsSmartMarkers";

            // -------------------------------------------------
            // 3. Prepare data source
            // -------------------------------------------------
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "HR" }
            };

            // -------------------------------------------------
            // 4. Set up WorkbookDesigner
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                LineByLine = false,          // use range smart markers
                UpdateReference = true       // ensure formulas referencing hidden sheets are updated
            };

            // Bind the data source to the name used in smart markers
            designer.SetDataSource("Employees", employees);

            // -------------------------------------------------
            // 5. Process all smart markers (including hidden sheet)
            // -------------------------------------------------
            designer.Process(); // processes every worksheet, hidden ones included

            // -------------------------------------------------
            // 6. Save the result
            // -------------------------------------------------
            workbook.Save("ProcessedWithHiddenSheet.xlsx");
        }
    }
}