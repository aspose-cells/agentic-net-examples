using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerMergeDemo
{
    // Simple data class used as a data source for smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Merge a range of cells (A1:C2) – this will be the placeholder area
            cells.Merge(0, 0, 2, 3); // rows: 2, columns: 3 (A1:C2)

            // 3. Put a smart marker placeholder inside the merged cell.
            // The placeholder follows the syntax "&=Data.Field"
            cells["A1"].PutValue("&=Data.Name");

            // 4. Prepare the data source that will replace the placeholder.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice Johnson", Age = 29 },
                new Person { Name = "Bob Smith", Age = 35 }
            };

            // 5. Set up the WorkbookDesigner, assign the workbook and the data source.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Data", people);

            // 6. Process the smart markers. The merged cell will retain its merge settings.
            designer.Process();

            // 7. Save the resulting workbook.
            workbook.Save("SmartMarkerMergedOutput.xlsx");
        }
    }
}