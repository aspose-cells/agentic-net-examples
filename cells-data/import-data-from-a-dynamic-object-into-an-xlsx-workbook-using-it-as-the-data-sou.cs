using System;
using System.Collections.Generic;
using System.Dynamic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Insert smart markers that will be replaced by the dynamic data
        // &=$Person.Name and &=$Person.Age refer to the data source named "Person"
        cells["A1"].PutValue("&=$Person.Name");
        cells["B1"].PutValue("&=$Person.Age");

        // Build a dynamic data source (list of ExpandoObject)
        var persons = new List<dynamic>();

        dynamic person1 = new ExpandoObject();
        person1.Name = "John Doe";
        person1.Age = 30;
        persons.Add(person1);

        dynamic person2 = new ExpandoObject();
        person2.Name = "Jane Smith";
        person2.Age = 28;
        persons.Add(person2);

        // Bind the dynamic collection to the smart marker name "Person"
        // (lifecycle rule: use WorkbookDesigner to set data source)
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;
        designer.SetDataSource("Person", persons);
        designer.Process(); // Populate the workbook with the dynamic data

        // Save the workbook to an XLSX file (lifecycle rule: save)
        workbook.Save("DynamicDataImport.xlsx");
    }
}