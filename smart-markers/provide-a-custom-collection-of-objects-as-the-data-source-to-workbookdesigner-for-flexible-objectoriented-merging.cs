// Title: Bind a Custom List<Person> to WorkbookDesigner Smart Markers in Aspose.Cells for .NET
// Description: Shows how to create a workbook, place smart markers (&Person.Name, &Person.Age), bind a List<Person> with WorkbookDesigner.SetDataSource, process the markers, and save the populated file as CustomCollectionOutput.xlsx.
// Keywords: Aspose.Cells | WorkbookDesigner | C# | .NET | smart markers | custom collection | SetDataSource | List<Person> | Excel export | object‑oriented merging
// Common Searches: Aspose.Cells bind List to smart markers | WorkbookDesigner custom object source example | C# set data source for smart markers | How to use a custom collection with WorkbookDesigner | Export List<Person> to Excel using Aspose.Cells
// Developer Intent: Populate an Excel worksheet from a custom collection by linking it to smart markers with WorkbookDesigner.
// Use Cases: Generate an employee roster where each row reflects a Person object. | Create a customer age report by merging a List<Person> into a template workbook. | Export in‑memory data structures to Excel without writing cell‑by‑cell code.
// AI Prompts: Add a bold header row and apply column width auto‑fit to the smart‑marker columns in the example. | Show how to bind multiple collections (e.g., List<Person> and List<Department>) to separate smart‑marker groups within the same workbook. | Explain how to use an ObservableCollection<Person> with WorkbookDesigner so that changes in the collection can be re‑processed to update the Excel file.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Shows how to create a workbook, place smart markers (&Person.Name, &Person.Age), bind a List<Person> with WorkbookDesigner.SetDataSource, process the markers, and save the populated file as CustomCollectionOutput.xlsx.
public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }

    public Person(string name, int age)
    {
        Name = name;
        Age = age;
    }
}

public class WorkbookDesignerCustomCollectionDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert smart markers that refer to the data source name "Person"
        sheet.Cells["A1"].PutValue("&Person.Name");
        sheet.Cells["B1"].PutValue("&Person.Age");

        // Initialize WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Prepare a custom collection of Person objects
        List<Person> persons = new List<Person>
        {
            new Person("John", 28),
            new Person("Emily", 34),
            new Person("Michael", 45)
        };

        // Bind the collection to the smart marker name "Person"
        designer.SetDataSource("Person", persons);

        // Process the smart markers and populate the worksheet
        designer.Process();

        // Save the populated workbook
        workbook.Save("CustomCollectionOutput.xlsx");
    }
}
