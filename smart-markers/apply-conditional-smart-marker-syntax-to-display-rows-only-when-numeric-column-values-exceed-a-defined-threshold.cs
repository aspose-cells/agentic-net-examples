using System;
using System.Data;
using Aspose.Cells;

class ConditionalSmartMarkerExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header cells
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Amount");

            // Insert smart markers in the template row (row 2)
            // Product name is always displayed
            sheet.Cells["A2"].PutValue("&=Products.ProductName");
            // Amount is displayed only when it exceeds the threshold (e.g., 100)
            sheet.Cells["B2"].PutValue("&IF(&=Amount>100, &=$Amount)");

            // Define the smart marker range and give it the required name
            Aspose.Cells.Range smartMarkerRange = sheet.Cells.CreateRange("A2:B2");
            smartMarkerRange.Name = "_CellsSmartMarkers";

            // Prepare the data source (DataTable) with sample data
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("ProductName", typeof(string));
            dt.Columns.Add("Amount", typeof(double));

            dt.Rows.Add("Item A", 80);   // Below threshold – row will be hidden (empty cells)
            dt.Rows.Add("Item B", 150);  // Above threshold – row will be shown
            dt.Rows.Add("Item C", 120);  // Above threshold – row will be shown
            dt.Rows.Add("Item D", 60);   // Below threshold – row will be hidden

            // Set up the WorkbookDesigner, assign the data source, and process only the smart marker range
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource(dt);
            // Process the defined range; true = preserve unrecognized markers
            designer.Process(smartMarkerRange, true);

            // Save the resulting workbook
            string outputPath = "ConditionalSmartMarkerOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}