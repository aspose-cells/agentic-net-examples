using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerMissingFieldHandling
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add smart markers that reference fields "Name" and "Age"
            // "Age" does NOT exist in the data source to simulate missing field
            sheet.Cells["A1"].PutValue("&=$Person.Name");
            sheet.Cells["A2"].PutValue("&=$Person.Age"); // missing field

            // Prepare data source with only the "Name" column
            DataTable dt = new DataTable("Person");
            dt.Columns.Add("Name", typeof(string));
            dt.Rows.Add("John Doe");

            // Set up the WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);

            // Process smart markers with error handling for missing fields
            try
            {
                designer.Process();
                Console.WriteLine("Smart markers processed successfully.");
            }
            catch (Exception ex)
            {
                // Handle the case where a smart marker references a missing field
                Console.WriteLine($"Error processing smart markers: {ex.Message}");
            }

            // Save the result (even if processing failed, the workbook can still be saved)
            workbook.Save("SmartMarkerResult.xlsx");
        }
    }
}