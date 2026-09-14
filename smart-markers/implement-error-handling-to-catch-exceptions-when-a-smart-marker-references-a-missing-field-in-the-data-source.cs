// Title: Use try‑catch to handle missing column errors in Aspose.Cells smart markers (C#)
// AI Prompts: Generate C# code that binds a DataTable to a smart marker and wraps WorkbookDesigner.Process in a try‑catch to capture missing field exceptions. | Show how to log the exception message when a smart marker references a column that does not exist in the data source. | Demonstrate saving the workbook after handling the error so the original smart marker text remains in the file.
// Common Searches: Aspose.Cells C# smart marker throws exception when column is missing | how to catch missing field error in WorkbookDesigner.Process | example of error handling for absent data columns in Aspose.Cells smart markers
// Tags: smart marker missing column exception Aspose.Cells | WorkbookDesigner.Process error handling C# | catch smart marker field not found Aspose.Cells | exception handling for Aspose.Cells smart markers | save workbook after smart marker error C#

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerErrorHandling
{
    // Demonstrates handling of missing field errors in smart markers
    // The example creates a workbook with a smart marker that references a non‑existent "Name" column, binds a DataTable containing only an "Age" column, and processes the markers inside a try‑catch block. The caught exception is logged, and the workbook is saved with the original smart marker text preserved.
    public class MissingFieldHandler
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that refers to a non‑existent column "Name"
            // The data source will only contain the column "Age"
            sheet.Cells["A1"].PutValue("&=$Employees.Name");

            // Prepare a DataTable with only the "Age" column (no "Name" column)
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add(30);
            dt.Rows.Add(45);

            // Set up the WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Bind the data source to the name used in the smart marker
            designer.SetDataSource("Employees", dt);

            // Process the smart markers with error handling
            try
            {
                // This will throw because the "Name" field is missing in the data source
                designer.Process();
                Console.WriteLine("Smart markers processed successfully.");
            }
            catch (Exception ex)
            {
                // Catch and display the error caused by the missing field
                Console.WriteLine($"Error processing smart markers: {ex.Message}");
            }

            // Save the workbook (the file will contain the original smart marker text)
            workbook.Save("MissingFieldResult.xlsx");
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            MissingFieldHandler.Run();
        }
    }
}
