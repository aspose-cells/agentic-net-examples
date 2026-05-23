using System;
using Aspose.Cells;
using System.Collections.Generic;

namespace AsposeCellsSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Insert smart markers with conditional logic
            //    &=$Person.Name  -> simple placeholder for Name
            //    &IF($Person.Age>30,\"Senior\",\"Junior\") -> conditional placeholder for Age
            sheet.Cells["A1"].PutValue("&=$Person.Name");
            sheet.Cells["B1"].PutValue("&IF($Person.Age>30,\"Senior\",\"Junior\")");

            // 4. Prepare anonymous objects as data source
            var persons = new List<object>
            {
                new { Name = "John Doe", Age = 35 },
                new { Name = "Jane Smith", Age = 28 }
            };

            // 5. Initialize WorkbookDesigner and bind the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // Bind the anonymous collection to the smart marker variable "Person"
            designer.SetDataSource("Person", persons);

            // 6. Process the smart markers (populate data)
            designer.Process();

            // 7. Save the result (lifecycle: save)
            workbook.Save("SmartMarkerConditionalOutput.xlsx");
        }
    }
}