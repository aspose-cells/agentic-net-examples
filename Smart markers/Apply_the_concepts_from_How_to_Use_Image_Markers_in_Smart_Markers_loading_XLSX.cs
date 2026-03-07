using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageMarkerDemo
{
    class Program
    {
        static void Main()
        {
            const string templatePath = "TemplateWithImageMarkers.xlsx";

            // Create a template workbook with image smart markers if it does not exist.
            if (!File.Exists(templatePath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                // Add a text smart marker for the name.
                ws.Cells["A1"].PutValue("&=Employees.Name");
                // Add an image smart marker for the photo.
                ws.Cells["B1"].PutValue("&=Employees.Photo");
                wb.Save(templatePath);
            }

            // Load the template workbook.
            var workbook = new Workbook(templatePath);

            // Prepare a list of employees with image data.
            var employees = new List<Employee>
            {
                new Employee
                {
                    Name = "John Doe",
                    Photo = GetImageBytes("john.jpg")
                },
                new Employee
                {
                    Name = "Jane Smith",
                    Photo = GetImageBytes("jane.jpg")
                }
            };

            // Bind the data source and process smart markers.
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", employees);
            designer.Process();

            // Save the result.
            workbook.Save("ResultWithImages.xlsx");
        }

        private static byte[] GetImageBytes(string path)
        {
            if (File.Exists(path))
            {
                return File.ReadAllBytes(path);
            }

            // Return a simple 1x1 PNG placeholder image.
            const string placeholderBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X0ZcAAAAASUVORK5CYII=";
            return Convert.FromBase64String(placeholderBase64);
        }
    }

    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public byte[] Photo { get; set; } = Array.Empty<byte>();
    }
}