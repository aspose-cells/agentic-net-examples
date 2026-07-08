using Aspose.Cells;
using System;
using System.Collections.Generic;

class VariableLoggingDesigner
{
    private readonly WorkbookDesigner _designer;

    public VariableLoggingDesigner(Workbook workbook)
    {
        // Initialize WorkbookDesigner with the provided workbook
        _designer = new WorkbookDesigner(workbook);

        // Attach a callback to log each smart‑marker processing event
        _designer.CallBack = new SmartMarkerLogger();
    }

    // Wrapper that logs the variable name/value and forwards the call to SetDataSource
    public void SetVariable(string name, object value)
    {
        Console.WriteLine($"SetVariable called: Name = {name}, Value = {value}");
        _designer.SetDataSource(name, value);
    }

    public void Process()
    {
        _designer.Process();
    }

    public void Save(string filePath)
    {
        _designer.Workbook.Save(filePath);
    }
}

// Implementation of ISmartMarkerCallBack that logs processing details
class SmartMarkerLogger : ISmartMarkerCallBack
{
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        Console.WriteLine($"SmartMarker processed - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
    }
}

// Sample data class used as a data source
class Employee
{
    public string Name { get; set; }
}

// Demonstration of logging variable assignments and smart‑marker processing
class Program
{
    static void Main()
    {
        // Create a new workbook that will serve as the template
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a variable smart marker and a data smart marker
        sheet.Cells["A1"].PutValue("&=$MyVar");
        sheet.Cells["A2"].PutValue("&=Employees.Name");

        // Initialize the designer with logging capabilities
        var designer = new VariableLoggingDesigner(workbook);

        // Log and set a simple variable
        designer.SetVariable("MyVar", "VariableValue");

        // Log and set a collection data source
        var employees = new List<Employee>
        {
            new Employee { Name = "John" },
            new Employee { Name = "Jane" }
        };
        designer.SetVariable("Employees", employees);

        // Process all smart markers
        designer.Process();

        // Save the resulting workbook
        designer.Save("LoggedSmartMarkers.xlsx");
    }
}