using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header cells
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");

            // Smart marker row – will be repeated for each item in the collection
            sheet.Cells["A2"].PutValue("&=Employees.Name");
            sheet.Cells["B2"].PutValue("&=Employees.Age");

            // Define the range that contains the smart markers and give it the required name
            Aspose.Cells.Range smartMarkerRange = sheet.Cells.CreateRange("A2:B2");
            smartMarkerRange.Name = "_CellsSmartMarkers";

            // Empty data source – no rows should be added
            List<Employee> employees = new List<Employee>(); // intentionally empty

            // Set up the designer
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Bind the empty collection to the smart marker name
            designer.SetDataSource("Employees", employees);

            // Process only the defined range; true = preserve unrecognized markers (not needed here)
            designer.Process(smartMarkerRange, true);

            // Save the workbook – no extra rows will be created
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple data class used for the smart marker
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}