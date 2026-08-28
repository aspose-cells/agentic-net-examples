// Title: How to configure Aspose.Cells WorkbookDesigner to convert empty strings to blank cells when using smart markers in C#
// AI Prompts: Set WorkbookDesigner.UpdateEmptyStringAsNull = true before calling Process() so that empty string values from a DataSet become blank cells during smart marker merging. | Place smart markers like &=$Products.Name in a worksheet, supply a DataTable containing empty strings, and enable blank conversion with the UpdateEmptyStringAsNull property. | Generate an Excel file with Aspose.Cells where any empty string fields from the source data are rendered as empty cells rather than literal empty strings.
// Common Searches: Aspose.Cells WorkbookDesigner blank cells for empty string values in smart markers C# | UpdateEmptyStringAsNull property example for merging DataSet into Excel | Treat empty strings as null when using smart markers with Aspose.Cells | How to prevent empty string literals in Excel output using Aspose.Cells WorkbookDesigner | Smart marker merge ignore empty string columns Aspose.Cells .NET
// Tags: WorkbookDesigner UpdateEmptyStringAsNull | smart markers blank cell handling | Aspose.Cells merge empty string as null | C# export DataSet to Excel with blanks | Excel smart marker empty string conversion

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsWorkbookDesignerDemo
{
    // Demonstrates creating a DataSet with empty string values, inserting smart markers into a workbook, enabling WorkbookDesigner.UpdateEmptyStringAsNull to treat those strings as blanks, processing the merge, and saving the resulting Excel file.
    class Program
    {
        static void Main()
        {
            // Create a sample DataSet containing empty strings
            DataSet ds = new DataSet();
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("Name");
            dt.Columns.Add("Price");
            dt.Rows.Add("Laptop", "");          // Price is empty string
            dt.Rows.Add("", "999.99");          // Name is empty string
            ds.Tables.Add(dt);

            // Create a new workbook and add designer markers
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("&=$Products.Name");   // Merge string property
            sheet.Cells["B1"].PutValue("&=$Products.Price");  // Merge string property

            // Initialize WorkbookDesigner and configure it to treat empty strings as blanks
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                UpdateEmptyStringAsNull = true   // Key setting
            };

            // Set the data source and process the merge
            designer.SetDataSource(ds.Tables["Products"]);
            designer.Process();

            // Verify the merged values (optional console output)
            Console.WriteLine("A1 merged value: " + sheet.Cells["A1"].StringValue);
            Console.WriteLine("B1 merged value: " + (sheet.Cells["B1"].StringValue == string.Empty ? "(blank)" : sheet.Cells["B1"].StringValue));

            // Save the resulting workbook
            workbook.Save("MergedResult.xlsx");
        }
    }
}
