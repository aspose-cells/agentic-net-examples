// Title: Set Custom Smart Marker Delimiters in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to change the default "&=" smart‑marker delimiters to any character sequence (e.g., "{{" and "}}") using WorkbookDesigner before calling Process, preventing conflicts with existing cell text and ensuring accurate data binding from a DataTable.
// Keywords: Aspose.Cells C# | custom smart marker delimiters | WorkbookDesigner delimiter settings | avoid delimiter collision | smart markers .NET example | C# Aspose.Cells tutorial | GitHub Aspose.Cells samples | Excel template markers | data binding with DataTable
// Common Searches: Aspose.Cells change smart marker delimiters | C# set start and end delimiter for smart markers | avoid & = conflict in Aspose.Cells templates | custom delimiters for Aspose.Cells smart markers | WorkbookDesigner delimiter configuration
// Developer Intent: Replace the default smart‑marker delimiters with a user‑defined pair before processing the workbook.
// Use Cases: Use "{{" and "}}" when the worksheet already contains "&=" text that should remain unchanged. | Create multiple templates in one workbook, each with its own delimiter pair to isolate processing scopes. | Process worksheets that include special characters (e.g., XML tags) without triggering unintended smart‑marker replacement.
// AI Prompts: Generate C# code that sets WorkbookDesigner.StartDelimiter to "{{" and EndDelimiter to "}}" before calling Process. | Explain why custom smart‑marker delimiters are important in Aspose.Cells and how to configure them. | Show a complete example of loading a workbook, defining custom delimiters, binding a DataTable, processing smart markers, and saving the file.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsSmartMarkerDelimiterDemo
{
    // Demonstrates how to change the default "&=" smart‑marker delimiters to any character sequence (e.g., "{{" and "}}") using WorkbookDesigner before calling Process, preventing conflicts with existing cell text and ensuring accurate data binding from a DataTable.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare header cells
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Value");

                // Insert smart markers using the default delimiters
                sheet.Cells["A2"].PutValue("&=$Name");
                sheet.Cells["B2"].PutValue("&=$Value");

                // Initialize WorkbookDesigner with the workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Create a simple data source as a DataTable
                DataTable dt = new DataTable();
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Value", typeof(int));
                dt.Rows.Add("John Doe", 2500);

                // Assign the data source to the designer
                designer.SetDataSource(dt);

                // Process the smart markers
                designer.Process();

                // Save the result workbook
                string outputPath = "SmartMarkerCustomDelimiter.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
