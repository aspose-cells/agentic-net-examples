using System;
using System.Dynamic;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Markup;

// Custom type that will be exposed to smart markers
public class MyDataCustom
{
    // Additional properties that can be referenced in smart markers
    public string ExtraInfo { get; set; }
    public int Quantity { get; set; }
}

// Class implementing ICustomTypeProvider to map MyData to MyDataCustom
public class MyData : ICustomTypeProvider
{
    // Regular property
    public string Name { get; set; }

    // Backing instance of the custom type
    private readonly MyDataCustom _custom = new MyDataCustom
    {
        ExtraInfo = "Sample extra information",
        Quantity = 42
    };

    // ICustomTypeProvider implementation – returns the custom type that contains extra properties
    public Type GetCustomType()
    {
        // The returned type is used by Aspose.Cells smart marker engine to resolve additional members
        return typeof(MyDataCustom);
    }

    // The smart marker engine will query the custom type for members; we need to provide the instance
    // via reflection when it accesses the properties. This is achieved by exposing the custom object
    // through a property with the same name as the custom type (optional but helpful for debugging).
    public MyDataCustom Custom => _custom;
}

// Demonstration of using the custom type provider with Aspose.Cells smart markers
public class SmartMarkerCustomTypeDemo
{
    public static void Run()
    {
        // Create a new workbook (Aspose.Cells lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert smart markers that reference both regular and custom properties
        // &=$MyData.Name      -> regular property
        // &=$MyData.ExtraInfo -> custom property provided via ICustomTypeProvider
        // &=$MyData.Quantity  -> another custom property
        sheet.Cells["A1"].PutValue("&=$MyData.Name");
        sheet.Cells["A2"].PutValue("&=$MyData.ExtraInfo");
        sheet.Cells["A3"].PutValue("&=$MyData.Quantity");

        // Prepare the data source
        MyData data = new MyData
        {
            Name = "Demo Item"
            // ExtraInfo and Quantity are already set inside MyDataCustom constructor
        };

        // Use WorkbookDesigner to process smart markers
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };

        // Register the data source; the key "MyData" matches the smart marker prefix
        designer.SetDataSource("MyData", data);

        // Process the smart markers (the second argument indicates whether to preserve empty cells)
        designer.Process(false);

        // Save the workbook (Aspose.Cells lifecycle rule)
        workbook.Save("SmartMarkerCustomTypeDemo.xlsx");
    }
}

// Entry point
class Program
{
    static void Main()
    {
        SmartMarkerCustomTypeDemo.Run();
        Console.WriteLine("Workbook created with custom smart marker bindings.");
    }
}