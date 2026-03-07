using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageMarkerDemo
{
    public class Person
    {
        public string Name { get; set; }
        public byte[] Photo { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Determine template path
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var templatePath = Path.Combine(baseDir, "Template.xlsx");

            Workbook workbook;

            if (File.Exists(templatePath))
            {
                // Load existing template
                workbook = new Workbook(templatePath);
            }
            else
            {
                // Create a new workbook with smart markers
                workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                // Smart marker for Name
                sheet.Cells["A1"].PutValue("&=Persons.Name");
                // Smart marker for Photo (image)
                sheet.Cells["B1"].PutValue("&=Persons.Photo");
            }

            // Prepare data source
            var persons = new List<Person>();
            var imagePath = Path.Combine(baseDir, "photo.jpg");
            byte[] photoBytes = File.Exists(imagePath) ? File.ReadAllBytes(imagePath) : Array.Empty<byte>();

            persons.Add(new Person
            {
                Name = "John Doe",
                Photo = photoBytes
            });

            // Set data source and process smart markers
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Persons", persons);
            designer.Process();

            // Save result
            var resultPath = Path.Combine(baseDir, "Result.xlsx");
            workbook.Save(resultPath);
        }
    }
}