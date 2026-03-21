using System;
using System.Collections.Generic;
using Aspose.Cells;

public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
    public Person(string name, int age) { Name = name; Age = age; }
}

public class SmartMarkerSlicerDemo
{
    public static void Main()
    {
        // 1. Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // 2. Add column headers
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Age");

        // 3. Insert smart markers with a slicer expression that filters Age > 30
        cells["A2"].PutValue("&=$People[?Age>30].Name");
        cells["B2"].PutValue("&=$People[?Age>30].Age");

        // 4. Mark the range containing the smart markers as a smart‑marker range.
        var smRange = cells.CreateRange("A2:B2");
        smRange.Name = "_CellsSmartMarkers";

        // 5. Prepare the data source (a list of Person objects)
        List<Person> people = new List<Person>
        {
            new Person("John", 25),
            new Person("Alice", 32),
            new Person("Bob", 45),
            new Person("Eve", 28)
        };

        // 6. Set up the WorkbookDesigner, assign the workbook and the data source
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };
        designer.SetDataSource("People", people);

        // 7. Process the smart markers – this will populate only rows where Age > 30
        designer.Process();

        // 8. Save the resulting workbook
        workbook.Save("SmartMarkerSlicerOutput.xlsx");
    }
}