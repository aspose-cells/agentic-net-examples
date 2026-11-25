// Title: Insert blank rows between groups using Aspose.Cells smart markers with (group:normal,skip:1) in C#
// Common Searches: how to add an empty row between grouped records with Aspose.Cells smart markers C# | Aspose.Cells smart marker group:normal skip:1 example for Excel export | C# code to separate data groups with blank rows using WorkbookDesigner | using Aspose.Cells smart markers to create grouped sections with spacing | generate Excel file with grouped rows and blank lines via Aspose.Cells C#
// Tags: Aspose.Cells smart marker grouping blank rows | WorkbookDesigner group normal skip 1 | C# Excel export grouped data Aspose.Cells | smart marker start end syntax Aspose.Cells | insert blank rows between groups Excel C#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// This C# example demonstrates how to use Aspose.Cells smart markers to group data and automatically insert a blank row between each group. The code creates a DataTable, places the smart marker start expression '&=dt.Start(group:normal,skip:1)&' in cell A1, defines column headers, and ends the range with '&=dt.End&'. WorkbookDesigner processes the markers, producing an Excel workbook where each distinct group (e.g., Fruits, Vegetables) is separated by an empty row. The resulting file is saved as 'GroupedData.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Prepare sample data to be used with smart markers
            DataTable dt = new DataTable();
            dt.Columns.Add("Category");
            dt.Columns.Add("Product");
            dt.Columns.Add("Price", typeof(double));

            dt.Rows.Add("Fruits", "Apple", 1.2);
            dt.Rows.Add("Fruits", "Banana", 0.8);
            dt.Rows.Add("Vegetables", "Carrot", 0.5);
            dt.Rows.Add("Vegetables", "Tomato", 0.9);

            // Insert smart marker expressions with grouping.
            // (group:normal,skip:1) adds a blank row between each group.
            // The marker starts at A1 and ends after the data rows.
            sheet.Cells["A1"].PutValue("&=dt.Start(group:normal,skip:1)&"); // start marker with grouping
            sheet.Cells["A2"].PutValue("&=Category&");   // column header for Category
            sheet.Cells["B2"].PutValue("&=Product&");    // column header for Product
            sheet.Cells["C2"].PutValue("&=Price&");      // column header for Price
            sheet.Cells["A3"].PutValue("&=dt.End&");     // end marker

            // Use WorkbookDesigner (compatible with all Aspose.Cells versions) to process smart markers
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("dt", dt);
            designer.Process();

            // Define output file path
            string outputPath = "GroupedData.xlsx";

            // Save the workbook (lifecycle rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
