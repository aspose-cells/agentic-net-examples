// Title: Aspose.Cells for .NET: Use a Master Smart Marker to Duplicate Rows in C#
// Description: Demonstrates how to add a master smart marker (e.g., "&=MasterData.Name" and "&=MasterData.Age") to an Excel template, bind a List<MasterRecord> with WorkbookDesigner, process the markers to repeat the row for each record, and save the workbook as an .xlsx file using C#.
// Keywords: Aspose.Cells | C# | .NET | master smart marker | repeat rows | WorkbookDesigner | Excel export | list binding | smart markers tutorial | Aspose.Cells US | Aspose.Cells Europe | Aspose.Cells Asia
// Common Searches: Aspose.Cells master smart marker example C# | repeat Excel rows for each object Aspose.Cells | bind List<T> to smart marker Aspose.Cells .NET | WorkbookDesigner process smart markers tutorial | generate Excel report from collection using Aspose.Cells
// Developer Intent: Insert a master smart marker into an Excel template so that the row is automatically duplicated for every item in a bound data collection.
// Use Cases: Create an employee directory by binding a List<Employee> to a master smart marker and exporting to Excel. | Generate an invoice where each product line repeats using a master smart marker bound to a List<Product>. | Export customer contact lists by repeating rows for each Customer object via a master smart marker.
// AI Prompts: Write C# code that adds a master smart marker row, binds a List<YourClass> to it with WorkbookDesigner, processes the markers, and saves the workbook using Aspose.Cells. | Explain step‑by‑step how to set up column headers and a master smart marker for row repetition in Aspose.Cells for .NET. | Show how to use WorkbookDesigner to bind multiple data sources, including a master list, and generate a populated Excel report with smart markers.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Sample data class representing a master record
    // Demonstrates how to add a master smart marker (e.g., "&=MasterData.Name" and "&=MasterData.Age") to an Excel template, bind a List<MasterRecord> with WorkbookDesigner, process the markers to repeat the row for each record, and save the workbook as an .xlsx file using C#.
    public class MasterRecord
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Set up column headers
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");

            // 3. Insert a master smart marker row.
            //    The marker "&=MasterData.Name" will be repeated for each item in the data source named "MasterData".
            sheet.Cells["A2"].PutValue("&=MasterData.Name");
            sheet.Cells["B2"].PutValue("&=MasterData.Age");

            // 4. Prepare sample data (list of master records)
            List<MasterRecord> masterData = new List<MasterRecord>
            {
                new MasterRecord { Name = "John Doe", Age = 30 },
                new MasterRecord { Name = "Jane Smith", Age = 28 },
                new MasterRecord { Name = "Bob Johnson", Age = 45 }
            };

            // 5. Create a WorkbookDesigner, assign the workbook and set the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // Bind the list to the smart marker name "MasterData"
            designer.SetDataSource("MasterData", masterData);

            // 6. Process the smart markers – rows will be repeated for each master record
            designer.Process();

            // 7. Save the resulting workbook
            workbook.Save("MasterSmartMarkerResult.xlsx");
        }
    }
}
