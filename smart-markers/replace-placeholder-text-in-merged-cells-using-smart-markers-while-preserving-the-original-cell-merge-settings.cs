using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ReplacePlaceholderInMergedCells
{
    // Simple data class used as a smart marker data source
    public class Person
    {
        public string Name { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Merge a range of cells (A1:B2). The merge settings will be preserved after processing.
            cells.Merge(firstRow: 0, firstColumn: 0, totalRows: 2, totalColumns: 2);

            // 3. Insert a smart marker placeholder inside the merged cell.
            //    The placeholder follows the smart marker syntax: &=DataSourceName.FieldName
            cells["A1"].PutValue("&=Person.Name");

            // 4. Prepare the data source that will replace the placeholder.
            List<Person> data = new List<Person>
            {
                new Person { Name = "John Doe" }
            };

            // 5. Set up the WorkbookDesigner, assign the workbook, and bind the data source.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Person", data);

            // 6. Process the smart markers. The merged cell remains merged after processing.
            designer.Process();

            // 7. Save the resulting workbook.
            workbook.Save("MergedCellPlaceholderReplaced.xlsx");
        }
    }
}