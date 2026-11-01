// Title: C# guide to implementing ISmartMarkerCallBack for detailed record‑level logging during Aspose.Cells smart marker merges
// Common Searches: how to log each record processed by smart markers using ISmartMarkerCallBack in Aspose.Cells .NET | C# example of custom ISmartMarkerCallBack implementation for Excel smart marker merge | capture merge events for each data row with Aspose.Cells smart markers | record‑level callback for smart marker processing in a .NET workbook designer | debug smart marker merges by tracking record callbacks in Aspose.Cells
// Tags: ISmartMarkerCallBack implementation C# | smart marker merge event logging Aspose.Cells | record processing callback for Excel smart markers | custom merge logger Aspose.Cells .NET | WorkbookDesigner data binding with callback | Aspose.Cells smart marker detailed logging

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an optional Excel template, creates a list of Employee objects, binds it to a WorkbookDesigner, and runs the smart marker merge. By adding an ISmartMarkerCallBack implementation you can capture each record's processing event, enabling detailed merge logs before saving the result workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Load a template if it exists; otherwise create a new empty workbook.
            Workbook workbook;
            const string templatePath = "Template.xlsx";
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook(); // empty workbook
            }

            // Sample data source for smart markers
            var employees = new List<Employee>
            {
                new Employee { Name = "Alice", Age = 30 },
                new Employee { Name = "Bob",   Age = 25 }
            };

            // Initialize WorkbookDesigner and register the data source
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", employees);

            // Execute the smart marker merge
            designer.Process();

            // Save the merged workbook
            const string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during processing:");
            Console.WriteLine(ex.Message);
        }
    }
}

// Simple POCO class representing an employee for smart marker data binding
public class Employee
{
    public string Name { get; set; }
    public int Age { get; set; }
}
