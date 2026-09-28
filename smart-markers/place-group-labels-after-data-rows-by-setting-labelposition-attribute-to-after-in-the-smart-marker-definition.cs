// Title: Insert a group label after the grouped rows using LabelPosition=After with Aspose.Cells smart markers in C#
// AI Prompts: Generate C# code that adds a smart marker with GroupLabel and configures its label to appear after the grouped rows using the LabelPosition attribute. | Show how to bind a DataTable to WorkbookDesigner, process the smart marker, and produce an .xlsx where the group label follows the data rows.
// Common Searches: how to make group label appear after rows in Aspose.Cells smart markers C# | C# Aspose.Cells example using LabelPosition attribute to position group labels | smart marker GroupLabel placement after data rows with WorkbookDesigner | Aspose.Cells generate Excel with group label after grouped data using C#
// Tags: Aspose.Cells smart marker group label positioning | LabelPosition attribute after rows | C# WorkbookDesigner data source binding | Excel export grouped rows with smart markers | Aspose.Cells group label after rows example

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a workbook, adds sample category data, inserts a smart marker {GroupLabel:LabelPosition=After}, binds a DataTable as the data source, processes the smart marker with WorkbookDesigner, and saves the file as SmartMarkerGroupLabelAfter.xlsx.
class SmartMarkerGroupLabelAfter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data that will be used for grouping
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        // Insert a smart marker for the group label.
        // The attribute LabelPosition=After places the label after the data rows.
        sheet.Cells["B1"].PutValue("{GroupLabel:LabelPosition=After}");

        // Prepare a data source (DataTable) that matches the smart marker.
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Category", typeof(string));
        dt.Rows.Add("A");
        dt.Rows.Add("B");
        dt.Rows.Add("C");

        // Use WorkbookDesigner to process the smart marker with the data source.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process();

        // Save the resulting workbook.
        workbook.Save("SmartMarkerGroupLabelAfter.xlsx");
    }
}
