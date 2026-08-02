// Title: Aspose.Cells for .NET – Master Smart Marker to Duplicate Rows per Record (C#)
// Description: Demonstrates how to place a master smart marker in an Excel template, name the range "_CellsSmartMarkers", bind a List<MasterRecord> to the marker with WorkbookDesigner, process the markers to generate a row for each record, and save the workbook. Ideal for dynamic employee, sales, or department reports.
// Keywords: Aspose.Cells master smart marker | C# smart markers repeat rows | WorkbookDesigner set data source list | define _CellsSmartMarkers range | dynamic Excel rows Aspose.Cells | Aspose.Cells example GitHub | Excel template smart marker C# | master‑detail smart markers Aspose
// Common Searches: Aspose.Cells master smart marker repeat rows | C# generate Excel rows from list using smart markers | how to use _CellsSmartMarkers range in Aspose.Cells | bind List<T> to smart marker name Aspose.Cells | Aspose.Cells smart marker example GitHub
// Developer Intent: Create an Excel worksheet where a template row is automatically duplicated for each item in a collection using a master smart marker.
// Use Cases: Export an employee directory to Excel with one row per employee. | Generate a sales team roster where each salesperson appears on a separate row. | Produce a department roster worksheet from a collection of department members.
// AI Prompts: Show how to add a subtotal row after the generated master rows using Aspose.Cells smart markers. | Provide code that combines a master smart marker with nested detail smart markers for master‑detail data. | Explain how to set the smart‑marker range programmatically when template rows are non‑contiguous.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Sample data class representing a master record
    // Demonstrates how to place a master smart marker in an Excel template, name the range "_CellsSmartMarkers", bind a List<MasterRecord> to the marker with WorkbookDesigner, process the markers to generate a row for each record, and save the workbook. Ideal for dynamic employee, sales, or department reports.
    public class MasterRecord
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Department { get; set; } = string.Empty;
    }

    public class MasterSmartMarkerExample
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Set up header row
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");
                sheet.Cells["C1"].PutValue("Department");

                // 3. Set up a template row that will be repeated for each master record
                //    Smart markers are placed in the cells of the template row
                sheet.Cells["A2"].PutValue("&=$Name");
                sheet.Cells["B2"].PutValue("&=$Age");
                sheet.Cells["C2"].PutValue("&=$Department");

                // 4. Define the range that contains the template row and give it the special name
                //    "_CellsSmartMarkers" tells the designer that this range contains smart markers.
                Aspose.Cells.Range templateRange = sheet.Cells.CreateRange("A2:C2");
                templateRange.Name = "_CellsSmartMarkers";

                // 5. Prepare sample data – a list of master records
                List<MasterRecord> masters = new List<MasterRecord>
                {
                    new MasterRecord { Name = "John Doe", Age = 35, Department = "Sales" },
                    new MasterRecord { Name = "Jane Smith", Age = 28, Department = "Marketing" },
                    new MasterRecord { Name = "Bob Johnson", Age = 42, Department = "HR" }
                };

                // 6. Create a WorkbookDesigner, assign the workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // 7. Bind the data source to the name used in the smart markers ("Master")
                designer.SetDataSource("Master", masters);

                // 8. Process the smart markers – rows will be generated for each master record
                designer.Process();

                // 9. Save the resulting workbook
                string outputPath = "MasterSmartMarkerOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during smart marker processing: {ex.Message}");
            }
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            MasterSmartMarkerExample.Run();
            Console.WriteLine("Workbook with master smart marker has been created.");
        }
    }
}
