// Title: Aspose.Cells WorkbookDesigner: Convert Empty Strings to Blank Cells with UpdateEmptyStringAsNull (C#)
// Description: Demonstrates how to configure WorkbookDesigner in Aspose.Cells for .NET so that empty string values from a DataSet are written as blank cells during smart‑marker processing by setting the UpdateEmptyStringAsNull property to true.
// Keywords: Aspose.Cells WorkbookDesigner | UpdateEmptyStringAsNull | empty string to blank cell | smart markers C# | Aspose.Cells data source null handling | Excel export empty values | C# Aspose.Cells example
// Common Searches: WorkbookDesigner treat empty strings as blanks | Aspose.Cells UpdateEmptyStringAsNull property | smart markers ignore empty strings | C# convert empty string to null Aspose.Cells | export Excel with blank cells instead of empty strings
// Developer Intent: Enable WorkbookDesigner to replace empty string fields from the data source with true blank cells when processing smart markers.
// Use Cases: Generating invoices where missing product names or prices should appear as empty cells. | Creating financial or inventory reports from databases that contain empty string fields, ensuring the Excel output shows blanks. | Automating Excel exports with smart markers while preventing placeholder empty‑string text from appearing in cells.
// AI Prompts: Show how to set WorkbookDesigner.UpdateEmptyStringAsNull in Aspose.Cells to treat empty strings as blanks. | Provide a C# code snippet that merges a DataSet containing empty strings into a workbook using smart markers and saves the result. | Explain the effect of UpdateEmptyStringAsNull on cell values and how to verify blank cells after processing.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to configure WorkbookDesigner in Aspose.Cells for .NET so that empty string values from a DataSet are written as blank cells during smart‑marker processing by setting the UpdateEmptyStringAsNull property to true.
    class WorkbookDesignerEmptyStringAsBlankDemo
    {
        static void Main()
        {
            // Create a sample DataSet containing empty strings
            DataSet ds = new DataSet();
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("Name");
            dt.Columns.Add("Price");
            dt.Rows.Add("Laptop", "");          // Empty string in Price column
            dt.Rows.Add("", "999.99");          // Empty string in Name column
            ds.Tables.Add(dt);

            // Create a new workbook and add designer markers
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("&=$Products.Name");
            sheet.Cells["B1"].PutValue("&=$Products.Price");

            // Initialize WorkbookDesigner and configure it to treat empty strings as null (blank)
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                UpdateEmptyStringAsNull = true   // Key setting
            };

            // Set the data source and process the markers
            designer.SetDataSource(ds.Tables["Products"]);
            designer.Process();

            // Save the result to verify that empty strings are treated as blanks
            workbook.Save("WorkbookDesigner_EmptyStringAsBlank_Output.xlsx");

            // Optional: display cell values to console for quick verification
            Console.WriteLine("A1 value: '" + sheet.Cells["A1"].StringValue + "'");
            Console.WriteLine("B1 value: '" + sheet.Cells["B1"].StringValue + "'");
        }
    }
}
