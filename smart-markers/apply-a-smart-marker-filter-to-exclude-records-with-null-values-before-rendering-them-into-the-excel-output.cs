// Title: C# example: filter out DBNull rows from a DataTable before processing smart markers with Aspose.Cells
// Common Searches: c# filter datatable rows with DBNull before using smart markers in Aspose.Cells | aspocells workbookdesigner exclude null records from data source | how to remove empty fields when generating Excel via smart markers in C# | example of cleaning DataTable to avoid null values in smart marker processing | smart markers ignore rows containing null values using Aspose.Cells
// Tags: smart marker data source filtering | filter DBNull rows Aspose.Cells | WorkbookDesigner exclude null records | Excel generation without null values C# | DataTable cleanup for smart markers

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The code loads an Excel template (or creates a new workbook), builds a DataTable of employee data, removes any rows that contain DBNull values, assigns the cleaned table to a WorkbookDesigner, processes the smart markers, saves the resulting workbook, and disposes resources.
class SmartMarkerFilterExample
{
    static void Main()
    {
        const string templatePath = "TemplateWithSmartMarkers.xlsx";
        const string outputPath = "OutputFilteredSmartMarkers.xlsx";

        Workbook workbook = null;

        try
        {
            // Load the Excel template that contains smart markers if it exists,
            // otherwise create a new empty workbook.
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found. Creating a new workbook.");
                workbook = new Workbook(); // empty workbook as fallback
            }

            // Prepare the data source (DataTable) with sample data
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Department", typeof(string));

            // Add rows – some contain null values that should be excluded
            dt.Rows.Add("John Doe", 30, "Sales");
            dt.Rows.Add(DBNull.Value, 25, "Marketing");   // Name is null
            dt.Rows.Add("Jane Smith", DBNull.Value, "HR"); // Age is null
            dt.Rows.Add("Mike Brown", 45, "IT");

            // Apply filter: remove rows where any column has a null (DBNull) value
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row = dt.Rows[i];
                foreach (DataColumn col in dt.Columns)
                {
                    if (row.IsNull(col))
                    {
                        dt.Rows.RemoveAt(i);
                        break;
                    }
                }
            }

            // Process the workbook with the filtered data source using WorkbookDesigner (Smart Marker)
            try
            {
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during Smart Marker processing: " + ex.Message);
                throw; // rethrow to be caught by outer catch
            }

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            // Ensure resources are released
            if (workbook != null)
            {
                workbook.Dispose();
            }
        }
    }
}
