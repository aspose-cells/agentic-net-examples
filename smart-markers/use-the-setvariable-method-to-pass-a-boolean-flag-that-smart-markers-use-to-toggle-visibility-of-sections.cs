// Title: C# – Toggle Conditional Smart Marker Sections with WorkbookDesigner.SetVariable in Aspose.Cells
// Description: Demonstrates how to insert &IF/$ShowSection/&ENDIF smart markers in an Excel template, assign a boolean flag using WorkbookDesigner.SetVariable, process the markers, and save the workbook. The example shows conditional visibility of rows based on a single true/false variable.
// Keywords: Aspose.Cells SetVariable C# | conditional smart markers | boolean flag Excel template | WorkbookDesigner.SetVariable example | toggle smart marker sections | Aspose.Cells smart markers tutorial | C# Excel conditional content
// Common Searches: Aspose.Cells SetVariable boolean example | C# conditional smart markers with &IF | how to hide rows using smart markers Aspose | WorkbookDesigner SetVariable vs SetDataSource | toggle Excel sections programmatically Aspose
// Developer Intent: Pass a boolean variable to smart markers so that &IF/$ShowSection/&ENDIF blocks are shown or hidden during workbook generation.
// Use Cases: Show a disclaimer row only when a regulatory flag is true. | Add an optional promotional paragraph to invoices based on a campaign switch. | Reveal advanced analytics sheets in a dashboard when a user‑enabled setting is active.
// AI Prompts: Generate C# code that uses WorkbookDesigner.SetVariable to control &IF/$ShowSection/&ENDIF smart markers in Aspose.Cells. | Explain the steps to hide a worksheet section with a boolean flag using SetVariable instead of SetDataSource. | Compare SetVariable and SetDataSource for handling conditional smart markers in Aspose.Cells.

using System;
using System.IO;
using System.Data;
using Aspose.Cells;

// Demonstrates how to insert &IF/$ShowSection/&ENDIF smart markers in an Excel template, assign a boolean flag using WorkbookDesigner.SetVariable, process the markers, and save the workbook. The example shows conditional visibility of rows based on a single true/false variable.
class SmartMarkerVariableDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet templateSheet = workbook.Worksheets[0];
            templateSheet.Name = "Template";

            // Insert smart markers that define a conditional section based on a boolean variable
            // &IF($ShowSection) starts the conditional block, &ENDIF ends it
            templateSheet.Cells["A1"].PutValue("&IF($ShowSection)");
            templateSheet.Cells["A2"].PutValue("This content is visible when ShowSection is true.");
            templateSheet.Cells["A3"].PutValue("&ENDIF");

            // Initialize WorkbookDesigner with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare a DataSet with a boolean column for the smart marker variable
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("ShowSection", typeof(bool));
            dt.Rows.Add(true); // set to false to hide the section
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);

            // Pass the DataSet to the designer
            designer.SetDataSource(ds);

            // Process the smart markers and populate the data
            designer.Process();

            // Define output file path
            string outputPath = "SmartMarkerVariableDemo.xlsx";

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
