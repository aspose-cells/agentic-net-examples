using System;
using System.Collections.Generic;
using Aspose.Cells;

// Simple POCO class representing a data record
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

// Custom data source implementing ICellsDataTable (inherits ICustomDataSource internally)
public class PersonDataSource : ICellsDataTable
{
    private readonly List<Person> _persons;
    private int _currentRow = -1;

    public PersonDataSource(List<Person> persons)
    {
        _persons = persons;
    }

    // Indexer for row/column access
    public object this[int rowIndex, int columnIndex] =>
        columnIndex == 0 ? _persons[rowIndex].Name : (object)_persons[rowIndex].Age;

    // Indexer for row access
    public object this[int rowIndex] => _persons[rowIndex];

    // Indexer for column name access
    public object this[string columnName] =>
        columnName == "Name" ? (object)_persons[_currentRow].Name : _persons[_currentRow].Age;

    public int RowCount => _persons.Count;
    public int ColumnCount => 2;
    public int Count => _persons.Count;
    public string[] Columns => new[] { "Name", "Age" };

    public void BeforeFirst()
    {
        _currentRow = -1;
    }

    public bool Next()
    {
        _currentRow++;
        return _currentRow < _persons.Count;
    }
}

class Program
{
    static void Main()
    {
        // Load an existing XLSX template workbook
        Workbook workbook = new Workbook("Template.xlsx");

        // Initialize WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Prepare sample data
        var data = new List<Person>
        {
            new Person("John Doe", 30),
            new Person("Jane Smith", 25)
        };

        // Bind the custom data source to the designer using a data source name
        designer.SetDataSource("Person", new PersonDataSource(data));

        // Process smart markers in the workbook
        designer.Process();

        // Save the processed workbook
        designer.Workbook.Save("Result.xlsx");
    }
}