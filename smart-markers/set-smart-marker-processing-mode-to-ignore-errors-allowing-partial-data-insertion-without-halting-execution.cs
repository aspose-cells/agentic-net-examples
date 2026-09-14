// Title: How to configure Aspose.Cells WorkbookDesigner to ignore smart marker errors and preserve unprocessed markers in C#
// AI Prompts: Generate C# code that calls WorkbookDesigner.Process(true) to keep smart markers when the data source contains null values. | Show an example of using the isPreserved flag with Aspose.Cells smart markers to skip errors and still write partial data to an Excel workbook.
// Common Searches: Aspose.Cells C# ignore errors while processing smart markers | WorkbookDesigner Process true keep unprocessed smart markers | How to handle null values in Aspose.Cells smart marker data source | Partial data export with smart markers without throwing exception in C# | Set smart marker processing mode to preserve markers Aspose.Cells
// Tags: WorkbookDesigner Process isPreserved true | Aspose.Cells smart marker error handling | ignore null values in smart marker data source | preserve unprocessed smart markers Excel | partial data insertion with Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace Demo
{
    // The example creates a workbook, inserts smart markers, registers a data source that includes a null value, and calls WorkbookDesigner.Process(true) so that unprocessed markers are kept and no exception is thrown. The resulting file is saved as SmartMarkerIgnoreError.xlsx.
    public class SmartMarkerIgnoreErrorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert smart markers into the worksheet
                sheet.Cells["A1"].PutValue("&=Employees.Name");
                sheet.Cells["B1"].PutValue("&=Employees.Age");

                // Define a named range that contains the smart markers.
                // The name "_CellsSmartMarkers" tells Aspose.Cells to treat this range as a smart‑marker block.
                sheet.Cells.CreateRange("A1:B1").Name = "_CellsSmartMarkers";

                // Prepare a data source that intentionally contains a missing value.
                var employees = new List<dynamic>
                {
                    new { Name = "John Doe", Age = 30 },
                    new { Name = "Jane Smith", Age = (int?)null }   // Missing value – will be ignored
                };

                // Set up the WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Register the data source with the name used in the smart markers.
                designer.SetDataSource("Employees", employees);

                // Process the smart markers.
                // Passing 'true' for the isPreserved parameter tells the designer to keep any
                // unprocessed (or error‑prone) smart markers instead of throwing an exception.
                designer.Process(true);

                // Save the resulting workbook.
                workbook.Save("SmartMarkerIgnoreError.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            SmartMarkerIgnoreErrorDemo.Run();
        }
    }
}
