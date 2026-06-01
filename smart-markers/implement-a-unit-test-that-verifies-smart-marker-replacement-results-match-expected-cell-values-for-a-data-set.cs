using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsSmartMarkerDemo
{
    // Simple data class used as data source for smart markers
    public class Employee
    {
        public string Name { get; set; } = string.Empty;   // Initialize to satisfy non‑nullable warning
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Insert smart markers into the template
                // &= Employees.Name  -> will be replaced by employee name
                // &= Employees.Age   -> will be replaced by employee age
                sheet.Cells["A1"].PutValue("&=Employees.Name");
                sheet.Cells["B1"].PutValue("&=Employees.Age");

                // Define the range that contains the smart markers and name it as required by Aspose.Cells
                AsposeRange markerRange = sheet.Cells.CreateRange("A1:B1");
                markerRange.Name = "_CellsSmartMarkers";

                // Prepare data source
                var employees = new List<Employee>
                {
                    new Employee { Name = "John", Age = 30 },
                    new Employee { Name = "Jane", Age = 25 }
                };

                // Set up the designer, bind the data source and process the smart markers
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Employees", employees);
                designer.Process();

                // Verify that the smart markers were replaced correctly
                bool success =
                    sheet.Cells["A1"].StringValue == "John" &&
                    sheet.Cells["B1"].IntValue == 30 &&
                    sheet.Cells["A2"].StringValue == "Jane" &&
                    sheet.Cells["B2"].IntValue == 25;

                Console.WriteLine(success
                    ? "Smart marker replacement succeeded."
                    : "Smart marker replacement failed.");

                // Optionally, save the workbook to inspect the result
                // string outputPath = "SmartMarkerResult.xlsx";
                // if (!File.Exists(Path.GetDirectoryName(outputPath)))
                // {
                //     Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
                // }
                // workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}