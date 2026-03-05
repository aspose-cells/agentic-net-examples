using System;
using System.Collections.Generic;
using System.Dynamic;
using Aspose.Cells;
using Aspose.Cells;

namespace AsposeCellsDynamicImportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a list of dynamic objects (ExpandoObject) as the data source
            var dynamicData = new List<dynamic>();

            dynamic item1 = new ExpandoObject();
            item1.Name = "Alice";
            item1.Age = 30;
            item1.City = "New York";
            dynamicData.Add(item1);

            dynamic item2 = new ExpandoObject();
            item2.Name = "Bob";
            item2.Age = 25;
            item2.City = "Chicago";
            dynamicData.Add(item2);

            dynamic item3 = new ExpandoObject();
            item3.Name = "Charlie";
            item3.Age = 35;
            item3.City = "Los Angeles";
            dynamicData.Add(item3);

            // Create a new workbook and a designer
            Workbook workbook = new Workbook();
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Insert smart markers into the first worksheet where data will be placed
            // &=$People.Name will be replaced by the Name property of each dynamic object
            Worksheet sheet = designer.Workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("&=$People.Name");
            sheet.Cells["B1"].PutValue("&=$People.Age");
            sheet.Cells["C1"].PutValue("&=$People.City");

            // Bind the dynamic list to the smart marker name "People"
            designer.SetDataSource("People", dynamicData);

            // Process the smart markers and generate the final worksheet
            designer.Process();

            // Save the workbook in XLSX format
            workbook.Save("DynamicDataImportDemo.xlsx");
        }
    }
}