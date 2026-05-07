using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ImageSmartMarkerDemo
{
    public class Employee
    {
        public string Name { get; set; } = "";
        public byte[] Photo { get; set; } = Array.Empty<byte>();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a workbook and add a worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Employees";

            // Set up smart markers for a table
            sheet.Cells["A1"].PutValue("&TableStart:Employees");
            sheet.Cells["A2"].PutValue("Name");
            sheet.Cells["B2"].PutValue("&=Name");
            sheet.Cells["C2"].PutValue("&=$Photo");
            sheet.Cells["A3"].PutValue("&TableEnd:Employees");

            // Prepare sample data
            List<Employee> employees = new List<Employee>();

            // A 1x1 red PNG image (base64 encoded)
            string base64RedPixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64RedPixel);

            employees.Add(new Employee
            {
                Name = "John Doe",
                Photo = imageBytes
            });

            // Initialize WorkbookDesigner and assign the workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Set the data source
            designer.SetDataSource("Employees", employees);

            // Process the smart markers
            designer.Process();

            // Save the result workbook
            string outputPath = "ResultWithEmbeddedImages.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'. Images have been embedded via smart markers.");
        }
    }
}