using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Header row
            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Age");
            ws.Cells["C1"].PutValue("Department");

            // Master smart marker row – this row will be repeated for each master record
            ws.Cells["A2"].PutValue("&=MasterData.Name");
            ws.Cells["B2"].PutValue("&=MasterData.Age");
            ws.Cells["C2"].PutValue("&=MasterData.Department");

            // Define the range that contains the smart markers.
            // When LineByLine is false, the range must be named "_CellsSmartMarkers".
            AsposeRange smRange = ws.Cells.CreateRange("A1:C2");
            smRange.Name = "_CellsSmartMarkers";

            // Sample data source – a list of master records
            var employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 35, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "HR" },
                new Employee { Name = "Bob Johnson", Age = 42, Department = "IT" }
            };

            // Set up the WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb
                // LineByLine is obsolete; using range smart markers instead
            };

            // Bind the data source to the name used in the smart markers
            designer.SetDataSource("MasterData", employees);

            // Process the template – rows will be repeated for each employee
            designer.Process();

            // Save the resulting workbook
            string outputPath = "MasterSmartMarkerOutput.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple data class representing a master record
    public class Employee
    {
        public string Name { get; set; } = null!;
        public int Age { get; set; }
        public string Department { get; set; } = null!;
    }
}