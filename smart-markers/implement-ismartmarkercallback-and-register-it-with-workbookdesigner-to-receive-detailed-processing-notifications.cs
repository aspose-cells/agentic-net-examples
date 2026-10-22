// Title: How to implement ISmartMarkerCallBack and attach it to WorkbookDesigner for detailed smart‑marker processing in Aspose.Cells for .NET
// AI Prompts: Create a C# class that implements Aspose.Cells.ISmartMarkerCallBack, overriding OnSmartMarkerStart, OnSmartMarkerEnd, and OnSmartMarkerError to log marker name, cell address, and any exception details. | Update the WorkbookDesigner example to instantiate the custom callback, assign it via designer.SmartMarkerCallBack, and then call designer.Process() so that each smart marker's processing events are captured and displayed.
// Common Searches: asp.net example of ISmartMarkerCallBack implementation with WorkbookDesigner | how to log smart marker processing steps using Aspose.Cells callback | register custom smart marker callback to capture row and column indices in C# | Aspose.Cells .NET smart marker event handling tutorial | debug smart markers with ISmartMarkerCallBack in Aspose.Cells
// Tags: custom smart marker callback C# | WorkbookDesigner smart marker processing | Aspose.Cells smart marker event handling | log smart marker details .NET | smart marker debugging with callback

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Sample data class used as a data source for smart markers
// The example loads a template workbook, sets a List<Person> as the data source for the "People" smart marker, registers a custom ISmartMarkerCallBack with WorkbookDesigner to receive start, end, and error notifications for each marker, processes the smart markers, and saves the resulting workbook while handling missing files and runtime exceptions.
public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
}

class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists before attempting to load it
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file '{templatePath}' not found.");
                return;
            }

            // Load the template workbook into a WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook(templatePath);

            // Prepare a data source for the smart marker "People"
            List<Person> people = new List<Person>
            {
                new Person { Name = "John", Age = 30 },
                new Person { Name = "Jane", Age = 25 }
            };
            designer.SetDataSource("People", people);

            // Process all smart markers in the workbook
            designer.Process();

            // Save the processed workbook
            designer.Workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
