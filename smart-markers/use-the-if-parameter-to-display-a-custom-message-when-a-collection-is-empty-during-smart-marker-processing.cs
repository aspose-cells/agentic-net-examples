// Title: C# – Show a Custom Message for an Empty Collection with the &If Smart Marker in Aspose.Cells
// Description: Illustrates the use of the &If parameter in Aspose.Cells smart markers to output a predefined text (e.g., "No employees found") when the bound collection is empty. The sample builds an in‑memory workbook, marks a range with smart markers, binds an empty List<Employee>, runs WorkbookDesigner.Process(), and saves the file.
// Keywords: Aspose.Cells | smart markers | &If parameter | empty collection | C# | WorkbookDesigner | SetDataSource | custom placeholder text | Excel report generation | no data message
// Common Searches: Aspose.Cells &If smart marker empty list | C# display message when smart marker collection is empty | How to use &If parameter in Aspose.Cells | Show "No data" placeholder with Aspose.Cells smart markers | WorkbookDesigner empty collection handling
// Developer Intent: Add an &If smart‑marker to an Excel template so that a specific message appears automatically when the associated data source contains zero rows.
// Use Cases: Insert "No employees found" in a HR report when the employee list is empty. | Show "No sales data available" in a financial dashboard if the sales collection returns no records. | Display "No records" in a summary sheet when a database query yields no rows.
// AI Prompts: Write C# code that uses Aspose.Cells smart markers with the &If parameter to display a custom message for an empty collection. | Explain the syntax and behavior of the &If smart marker in Aspose.Cells and how it interacts with WorkbookDesigner.SetDataSource. | Provide step‑by‑step guidance to modify an existing Excel template to include an &If smart marker for handling zero‑row data sources.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerIfDemo
{
    // Simple data class for demonstration
    // Illustrates the use of the &If parameter in Aspose.Cells smart markers to output a predefined text (e.g., "No employees found") when the bound collection is empty. The sample builds an in‑memory workbook, marks a range with smart markers, binds an empty List<Employee>, runs WorkbookDesigner.Process(), and saves the file.
    public class Employee
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (template) in memory
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Add smart markers
                //    The first marker uses the If parameter to display a custom message
                //    when the collection "Employees" is empty.
                //    Syntax: &If=CollectionName?Message
                cells["A1"].PutValue("&If=Employees?No employees found");

                //    Define the range that will be processed for the collection items.
                //    These markers will be repeated for each item in the collection.
                cells["A2"].PutValue("&=Employees.Name");
                cells["B2"].PutValue("&=Employees.Age");

                //    Name the range that contains the smart markers.
                //    The special name "_CellsSmartMarkers" tells the designer to process this range.
                Aspose.Cells.Range smartRange = cells.CreateRange("A1:B2");
                smartRange.Name = "_CellsSmartMarkers";

                // 3. Prepare an empty data source (no employees)
                List<Employee> employees = new List<Employee>(); // empty collection

                // 4. Set up the WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                //    Bind the empty collection to the smart marker name "Employees"
                designer.SetDataSource("Employees", employees);

                // 5. Process the smart markers
                //    The If parameter will cause the message "No employees found" to appear
                //    because the collection is empty.
                designer.Process();

                // 6. Save the result to a file
                string outputPath = "SmartMarkerIfResult.xlsx";
                designer.Workbook.Save(outputPath);

                // Optional: output a confirmation
                Console.WriteLine($"Workbook saved. Check '{outputPath}' for the custom empty‑collection message.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
