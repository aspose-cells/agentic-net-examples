using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerRangeDemo
{
    public class Person
    {
        public string Name { get; set; } = null!;
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a template workbook ----------
            Workbook templateWb = new Workbook();
            Worksheet sheet = templateWb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");

            // Smart marker row – will be repeated for each item in the data source
            cells["A2"].PutValue("&=$Data.Name");
            cells["B2"].PutValue("&=$Data.Age");

            // Define the smart marker range block (A2:B2) and give it the required name
            Aspose.Cells.Range smartRange = cells.CreateRange("A2:B2");
            smartRange.Name = "_CellsSmartMarkers";

            // ---------- Prepare variable‑length data ----------
            List<Person> people = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 28 },
                new Person { Name = "Bob Johnson", Age = 45 }
            };

            // ---------- Bind data to the smart marker range ----------
            WorkbookDesigner designer = new WorkbookDesigner(templateWb);
            designer.SetDataSource("Data", people);
            designer.Process();

            // ---------- Save the result ----------
            designer.Workbook.Save("SmartMarkerRangeResult.xlsx");
        }
    }
}