// Title: Aspose.Cells WorkbookDesigner: Convert Empty Strings to Blank Cells with UpdateEmptyStringAsNull (C#)
// Description: Demonstrates how to set WorkbookDesigner.UpdateEmptyStringAsNull to true so that empty string values from a DataSet/DataTable are merged as null (blank) cells when processing smart markers. The example creates a workbook, adds markers, binds the data source, runs the designer, and saves the result.
// Keywords: Aspose.Cells | WorkbookDesigner | UpdateEmptyStringAsNull | empty string handling | blank cells | smart markers | C# | .NET | Excel export | DataSet | DataTable | null conversion
// Common Searches: Aspose.Cells UpdateEmptyStringAsNull example | WorkbookDesigner treat empty strings as null | smart markers ignore empty values C# | convert empty string to blank cell Aspose.Cells | how to hide empty strings in Excel export using Aspose
// Developer Intent: Configure WorkbookDesigner to replace empty string fields from a data source with null/blank cells during smart‑marker processing.
// Use Cases: Generating invoices where missing product names or prices should appear as empty cells rather than literal empty strings. | Creating sales or inventory reports from optional‑field DataTables while keeping the Excel layout clean. | Exporting a database view to Excel via smart markers, ensuring that nullable text columns render as blank cells.
// AI Prompts: Show C# code that sets WorkbookDesigner.UpdateEmptyStringAsNull to true for smart marker processing in Aspose.Cells. | Provide a step‑by‑step example of binding a DataSet with empty strings to WorkbookDesigner and saving the workbook with blank cells. | Explain how UpdateEmptyStringAsNull affects cell formatting and how to verify that empty strings become blank cells after processing.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to set WorkbookDesigner.UpdateEmptyStringAsNull to true so that empty string values from a DataSet/DataTable are merged as null (blank) cells when processing smart markers. The example creates a workbook, adds markers, binds the data source, runs the designer, and saves the result.
    public class WorkbookDesignerEmptyStringAsBlankDemo
    {
        public static void Run()
        {
            try
            {
                // Create a sample DataSet containing empty strings
                DataSet ds = new DataSet();
                DataTable dt = new DataTable("Products");
                dt.Columns.Add("Name");
                dt.Columns.Add("Price");
                dt.Rows.Add("Laptop", "");          // Price is empty
                dt.Rows.Add("", "999.99");          // Name is empty
                ds.Tables.Add(dt);

                // Create a new workbook and place designer markers
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("&=$Products.Name");
                sheet.Cells["B1"].PutValue("&=$Products.Price");

                // Initialize WorkbookDesigner and set it to treat empty strings as blanks (null)
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    UpdateEmptyStringAsNull = true // key property
                };

                // Bind the data source and process the markers
                designer.SetDataSource(ds.Tables["Products"]);
                designer.Process();

                // Save the resulting workbook
                string outputPath = "WorkbookDesigner_EmptyStringAsBlank_Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookDesignerEmptyStringAsBlankDemo.Run();
        }
    }
}
