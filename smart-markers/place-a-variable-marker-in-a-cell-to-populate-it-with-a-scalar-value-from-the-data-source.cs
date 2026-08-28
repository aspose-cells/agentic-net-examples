// Title: Place a scalar smart marker in a cell and populate it from a DataSet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, inserts the smart marker "&=Product.Name" into cell A1, binds a DataSet containing a Product table to WorkbookDesigner, processes the marker, and saves the file as an .xlsx document. | Modify the example to read product information from a JSON file, convert the JSON to a DataSet, and use a single‑value smart marker to write the first product's name into a designated worksheet cell. | Expand the sample so that a list of Product objects is converted to a DataSet and a smart‑marker range is used to fill an entire column with product names instead of a single cell.
// Common Searches: how to bind a DataSet to WorkbookDesigner for smart marker replacement in Aspose.Cells .NET | C# example of using & =Product.Name scalar smart marker in an Excel worksheet | populate a single Excel cell from a POCO object using Aspose.Cells smart markers | Aspose.Cells process smart markers after setting data source with DataSet | save workbook after smart marker processing in C#
// Tags: WorkbookDesigner SetDataSource DataSet C# | replace smart marker with object property value | Aspose.Cells generate Excel from POCO | save workbook after processing markers | C# example of smart marker in cell A1

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple data class containing a scalar property
    // Shows how to insert a scalar smart marker "&=Product.Name" into cell A1, create a DataSet with a Product table, bind it to WorkbookDesigner, process the marker, and save the resulting workbook as SmartMarkerOutput.xlsx.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Place a smart marker in cell A1.
                // The marker "&=Product.Name" tells Aspose.Cells to replace it with the value of the "Name" column
                // from the "Product" table that will be provided later.
                worksheet.Cells["A1"].PutValue("&=Product.Name");

                // Prepare the data source with a scalar value
                Product product = new Product { Name = "Gadget" };

                // Convert the product to a DataSet (required by WorkbookDesigner.SetDataSource)
                DataSet dataSet = new DataSet();
                DataTable table = new DataTable("Product");
                table.Columns.Add("Name", typeof(string));
                table.Rows.Add(product.Name);
                dataSet.Tables.Add(table);

                // Initialize WorkbookDesigner with the workbook (lifecycle: load)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Set the data source; the table name "Product" matches the smart marker.
                designer.SetDataSource(dataSet);

                // Process the smart markers and populate the cell (feature: smart marker processing)
                designer.Process();

                // Save the workbook (lifecycle: save)
                string outputPath = "SmartMarkerOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
