// Title: Use a DataTable as the data source and apply IF conditional logic in Aspose.Cells smart markers (C#)
// Common Searches: how to set a DataTable as the data source for WorkbookDesigner smart markers in C# | aspnet cells smart marker IF syntax for conditional values | example of conditional smart markers based on age using Aspose.Cells | populate Excel worksheet with smart markers from a .NET DataTable | using WorkbookDesigner to process smart markers with conditional expressions
// Tags: WorkbookDesigner DataTable source | smart markers IF expression Aspose.Cells | conditional smart markers C# | Excel generation from DataTable Aspose.Cells | smart marker data binding .NET

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds column headers, inserts smart markers with an IF expression to label ages as "Adult" or "Minor", builds a DataTable containing Name and Age rows, assigns the DataTable as the data source to a WorkbookDesigner, processes the smart markers, and saves the resulting Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up column headers
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("Category");

            // Insert smart markers starting from row 2
            sheet.Cells["A2"].PutValue("&=Name");
            sheet.Cells["B2"].PutValue("&=Age");
            sheet.Cells["C2"].PutValue("&=IF([Age]>=18,\"Adult\",\"Minor\")");

            // Prepare data source as a DataTable (required by WorkbookDesigner)
            DataTable peopleTable = new DataTable("People");
            peopleTable.Columns.Add("Name", typeof(string));
            peopleTable.Columns.Add("Age", typeof(int));

            peopleTable.Rows.Add("John", 30);
            peopleTable.Rows.Add("Alice", 15);
            peopleTable.Rows.Add("Bob", 22);

            // Process smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(peopleTable);
            designer.Process();

            // Save the workbook (lifecycle rule: save)
            string outputFile = "SmartMarkersConditional.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
