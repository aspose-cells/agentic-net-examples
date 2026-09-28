// Title: Display a custom message with Aspose.Cells IF smart marker when a bound collection is empty (C#)
// AI Prompts: Write C# code that inserts an &If smart marker into a worksheet cell to show a fallback text when the Employees list has no items. | Demonstrate how to bind an empty List<Employee> to WorkbookDesigner and process smart markers so the IF parameter writes the custom message. | Explain the syntax of the &If parameter for conditional output in Aspose.Cells smart markers.
// Common Searches: asp.net aspocells smart marker if parameter empty list example | c# aspose.cells show no data message when collection is empty | how to use &If in smart markers to display custom text for empty data source | workbookdesigner conditional smart marker output based on empty collection c#
// Tags: Aspose.Cells IF smart marker | conditional smart marker output | WorkbookDesigner empty collection handling | C# Excel generation with smart markers | custom message for no data in Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerIfDemo
{
    // Simple data class (not used in this example because the collection is empty)
    // The example creates a workbook, places an &If smart marker ("&If=Employees?\"No employees found\"") in cell A1, binds an empty List<Employee> as the data source, processes the markers with WorkbookDesigner, and saves the file, resulting in the custom message appearing when the collection is empty.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Insert a smart marker that uses the IF parameter.
            //    The syntax "&If=Employees?\"No employees found\"" means:
            //    - If the collection "Employees" is empty, display the text inside the quotes.
            //    - If the collection has items, the smart marker will be ignored (no output).
            sheet.Cells["A1"].PutValue("&If=Employees?\"No employees found\"");

            // 3. Prepare an empty data source (the collection is intentionally left empty)
            List<Employee> employees = new List<Employee>(); // empty list

            // 4. Set up the WorkbookDesigner, assign the workbook and the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Employees", employees);

            // 5. Process the smart markers. The IF parameter will cause the custom message
            //    to be written into cell A1 because the collection is empty.
            designer.Process();

            // 6. Save the result to a file
            workbook.Save("SmartMarkerIfResult.xlsx");

            Console.WriteLine("Processing complete. Check 'SmartMarkerIfResult.xlsx' for the custom message.");
        }
    }
}
