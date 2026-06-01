using Aspose.Cells;
using System;
using System.Data;
using System.IO;

class SetVariableDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Template";

            // Add a smart marker that references a boolean variable
            // This cell will display the value of the variable (True/False)
            sheet.Cells["A1"].PutValue("&=$ShowSection");

            // Add a placeholder that will be replaced based on the variable.
            // In a real template you would use conditional smart marker syntax.
            sheet.Cells["A2"].PutValue("&ShowSection?VisibleContent:HiddenContent");

            // Create a WorkbookDesigner instance for the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare a DataSet with a DataTable containing the boolean flag
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("ShowSection", typeof(bool));
            dt.Rows.Add(true); // Change to false to hide the section

            DataSet ds = new DataSet();
            ds.Tables.Add(dt);

            // Pass the data source to the designer
            designer.SetDataSource(ds);

            // Process all smart markers in the workbook
            designer.Process();

            // Define output file path
            string outputPath = "SetVariableDemo.xlsx";

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}