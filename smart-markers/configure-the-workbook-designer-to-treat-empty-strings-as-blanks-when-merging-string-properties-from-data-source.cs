using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWorkbookDesignerDemo
{
    class Program
    {
        static void Main()
        {
            try
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
                sheet.Cells["A1"].PutValue("&=$Products.Name");
                sheet.Cells["B1"].PutValue("&=$Products.Price");

                // Initialize WorkbookDesigner and configure it to treat empty strings as nulls
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    UpdateEmptyStringAsNull = true   // Key setting
                };

                // Set the data source and process the markers
                designer.SetDataSource(ds.Tables["Products"]);
                designer.Process();

                // Verify the result (optional console output)
                Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue); // Expected: "Laptop"

                // Use Cell.Value to determine if the cell is blank (null)
                var b1Cell = sheet.Cells["B1"];
                string b1Output = b1Cell.Value == null ? "Blank (null)" : b1Cell.StringValue;
                Console.WriteLine("B1 value: " + b1Output);

                // Save the workbook to a file
                string outputPath = "WorkbookDesigner_EmptyStringAsNull_Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}