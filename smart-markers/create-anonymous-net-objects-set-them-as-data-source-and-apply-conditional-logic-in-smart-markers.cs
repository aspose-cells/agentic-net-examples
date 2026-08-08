// Title: Aspose.Cells for .NET – Bind Anonymous Objects & Use &IF Conditional Smart Markers
// Description: Create a workbook, add smart markers (&=$Person.Name and &IF($Person.Age>30,"Senior","Junior")), bind an array of anonymous C# objects as the "Person" data source with WorkbookDesigner, process the template, and save the XLSX file.
// Keywords: Aspose.Cells | smart markers | anonymous objects | C# | .NET | &IF function | conditional smart marker | WorkbookDesigner | set data source | Excel template generation | dynamic Excel report
// Common Searches: asp.net bind anonymous type to Aspose.Cells smart markers | Aspose.Cells &IF smart marker example C# | set data source for WorkbookDesigner without POCO class | conditional smart marker syntax Aspose.Cells | how to use anonymous objects with Aspose.Cells templates
// Developer Intent: Bind an array of anonymous .NET objects to WorkbookDesigner and apply an &IF smart marker that outputs different text based on a property value.
// Use Cases: Generate a staff roster where each row shows the employee name and labels them as Senior or Junior according to age. | Produce a sales dashboard that marks each entry as "Target Met" or "Target Missed" using &IF logic on an anonymous sales collection. | Create an inventory sheet that flags items as "Low Stock" or "Sufficient" based on quantity values supplied by an anonymous data set.
// AI Prompts: Write C# code that creates an array of anonymous objects, sets it as the "Person" data source for WorkbookDesigner, and uses &IF smart markers to display "Senior" or "Junior" based on age. | Explain step‑by‑step how to apply the &IF function in Aspose.Cells smart markers when the source data is an anonymous type collection. | Generate a complete Aspose.Cells example that demonstrates conditional formatting with &IF and anonymous objects for a sales report.

using System;
using Aspose.Cells;

// Create a workbook, add smart markers (&=$Person.Name and &IF($Person.Age>30,"Senior","Junior")), bind an array of anonymous C# objects as the "Person" data source with WorkbookDesigner, process the template, and save the XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Insert smart markers:
        //   &=$Person.Name   -> simple data binding
        //   &IF($Person.Age>30,"Senior","Junior") -> conditional logic
        ws.Cells["A1"].PutValue("&=$Person.Name");
        ws.Cells["B1"].PutValue("&IF($Person.Age>30,\"Senior\",\"Junior\")");

        // Prepare anonymous objects as the data source
        var persons = new[]
        {
            new { Name = "John Doe", Age = 35 },
            new { Name = "Jane Smith", Age = 28 }
        };

        // Set up the designer, bind the data source and process the smart markers
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = wb;
        designer.SetDataSource("Person", persons);
        designer.Process();

        // Save the populated workbook
        wb.Save("SmartMarkerConditionalOutput.xlsx");
    }
}
