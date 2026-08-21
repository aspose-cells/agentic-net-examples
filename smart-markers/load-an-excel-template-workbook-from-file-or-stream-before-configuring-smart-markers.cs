// Title: Load an Excel template (file or stream) and process smart markers with Aspose.Cells for .NET
// Description: Demonstrates how to load a template workbook (Template.xlsx) from a file path or a MemoryStream, bind a List<Employee> to the WorkbookDesigner, process smart markers, and save the populated result as Result.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells load workbook from stream | smart markers template loading | WorkbookDesigner data source | process smart markers C# | save processed workbook Aspose.Cells | Excel template from MemoryStream | .NET Excel smart markers example
// Common Searches: load Excel template from MemoryStream Aspose.Cells | Aspose.Cells smart markers file vs stream | WorkbookDesigner load template workbook C# | how to bind List<Employee> to smart markers | process smart markers and save result
// Developer Intent: Load a template workbook (file or stream), bind data to smart markers, process them, and save the final Excel file.
// Use Cases: Generate reports from a pre‑designed Excel template stored on disk. | Read a template saved as a byte array in a database, populate it with employee data, and export the result. | Reuse the same loaded workbook with multiple data sets to create batch reports without re‑reading the template file.
// AI Prompts: Show C# code to load an Excel template from a MemoryStream and process smart markers with Aspose.Cells. | Give an example that reads a template workbook from a byte array, sets a List<Employee> as the smart marker data source, processes the markers, and saves the output. | Explain error handling best practices when loading a template workbook for smart markers in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple data class used as a smart marker data source
    // Demonstrates how to load a template workbook (Template.xlsx) from a file path or a MemoryStream, bind a List<Employee> to the WorkbookDesigner, process smart markers, and save the populated result as Result.xlsx using Aspose.Cells for .NET.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Load the Excel template workbook.
            //    The template file should contain smart markers like
            //    &Employee.Name and &Employee.Age in the worksheet.
            // ------------------------------------------------------------
            Workbook templateWorkbook = new Workbook("Template.xlsx"); // Load from file

            // If you prefer loading from a stream, uncomment the following lines:
            // byte[] fileBytes = File.ReadAllBytes("Template.xlsx");
            // using (MemoryStream ms = new MemoryStream(fileBytes))
            // {
            //     templateWorkbook = new Workbook(ms); // Load from stream
            // }

            // ------------------------------------------------------------
            // 2. Create a WorkbookDesigner and assign the loaded workbook.
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = templateWorkbook;

            // ------------------------------------------------------------
            // 3. Prepare the data source that will populate the smart markers.
            // ------------------------------------------------------------
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 30 },
                new Employee { Name = "Jane Smith", Age = 28 }
            };

            // Bind the data source to the smart marker name "Employee".
            designer.SetDataSource("Employee", employees);

            // ------------------------------------------------------------
            // 4. Process the smart markers to fill the worksheet with data.
            // ------------------------------------------------------------
            designer.Process();

            // ------------------------------------------------------------
            // 5. Save the processed workbook to a new file.
            // ------------------------------------------------------------
            designer.Workbook.Save("Result.xlsx", SaveFormat.Xlsx);
        }
    }
}
