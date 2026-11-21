// Title: Place a group label after data rows with LabelPosition='After' in Aspose.Cells smart markers (C#)
// Common Searches: C# Aspose.Cells smart marker labelposition after to put group header after rows | How to add a group label after data rows using smart markers in Aspose.Cells | Setting the LabelPosition attribute to After in WorkbookDesigner for smart markers | Aspose.Cells example for exporting a DataTable with a trailing group label
// Tags: group label placement after rows smart marker | configure smart marker label position in Aspose.Cells | Aspose.Cells WorkbookDesigner data grouping with smart markers | C# generate Excel using smart markers and DataTable

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads a template workbook (or creates a new one), defines a smart marker in cell A1 using '&=Employees.Name[LabelPosition=After]' to place the group label after the data rows, fills a DataTable with employee names, assigns it as the data source for the smart marker via WorkbookDesigner, processes the markers, and saves the result to 'Result.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the template workbook
            string templatePath = "Template.xlsx";

            // Load existing template or create a new workbook if the file is missing
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
                Console.WriteLine($"Template file not found. A new workbook has been created.");
            }

            // Define the smart marker with LabelPosition set to "After"
            // This places the group label after the data rows.
            Worksheet sheet = workbook.Worksheets[0];
            Cell smartMarkerCell = sheet.Cells["A1"];
            smartMarkerCell.PutValue("&=Employees.Name[LabelPosition=After]");

            // Prepare sample data source
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Rows.Add("Alice");
            dt.Rows.Add("Bob");
            dt.Rows.Add("Charlie");

            // Set the data source for the smart marker and process it
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", dt);
            designer.Process();

            // Save the resulting workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
