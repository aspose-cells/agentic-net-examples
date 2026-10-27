// Title: Map a C# DataTable to a predefined DetailTable smart marker in an Excel template using Aspose.Cells
// Common Searches: Aspose.Cells C# bind DataTable to smart marker table named DetailTable | populate predefined Excel table with WorkbookDesigner SetDataSource in .NET | example of filling a DetailTable smart marker from a DataTable using Aspose.Cells | C# code to map detail rows to smart marker placeholders in an Excel template | load or create workbook with smart markers and fill with DataTable Aspose.Cells
// Tags: WorkbookDesigner SetDataSource with DataTable | Aspose.Cells smart marker table binding | populate Excel smart marker DetailTable from C# DataTable | predefined table mapping using Aspose.Cells | Excel template smart markers .NET example

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // The example loads or creates an Excel workbook containing smart‑marker placeholders for a DetailTable, constructs a C# DataTable with Item, Quantity, and UnitPrice columns, adds sample rows, binds the DataTable to the smart marker named "DetailTable" via WorkbookDesigner.SetDataSource, processes the template to fill the predefined table, and saves the resulting workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Load the template workbook if it exists; otherwise create a simple workbook with smart‑marker placeholders.
                Workbook workbook;
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "Sheet1";

                    // Add smart‑marker placeholders for the DetailTable.
                    sheet.Cells["A1"].PutValue("&DetailTable.Item");
                    sheet.Cells["B1"].PutValue("&DetailTable.Quantity");
                    sheet.Cells["C1"].PutValue("&DetailTable.UnitPrice");
                }

                // Create a designer to work with smart markers.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Prepare the detail data as a DataTable.
                DataTable detailTable = new DataTable("DetailTable");
                detailTable.Columns.Add("Item", typeof(string));
                detailTable.Columns.Add("Quantity", typeof(int));
                detailTable.Columns.Add("UnitPrice", typeof(double));

                // Populate the DataTable with sample rows.
                detailTable.Rows.Add("Pen", 15, 1.25);
                detailTable.Rows.Add("Notebook", 8, 3.40);
                detailTable.Rows.Add("Eraser", 20, 0.55);

                // Bind the DataTable to the smart marker named "DetailTable".
                designer.SetDataSource("DetailTable", detailTable);

                // Process the template – data will be written into the predefined table.
                designer.Process();

                // Save the resulting workbook.
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
