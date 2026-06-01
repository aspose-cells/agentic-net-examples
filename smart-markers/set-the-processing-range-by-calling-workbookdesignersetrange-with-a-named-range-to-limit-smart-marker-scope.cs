using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerRangeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Insert sample data and smart markers
                // Header row
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");

                // Smart marker row (will be repeated for each data item)
                cells["A2"].PutValue("&=$Product");
                cells["B2"].PutValue("&=$Price");

                // 3. Define a named range that encloses the smart marker row
                // This range can be used to limit processing scope if needed
                Aspose.Cells.Range smartMarkerRange = cells.CreateRange("A2:B2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // 4. Prepare a simple data source (ArrayList of anonymous objects)
                ArrayList data = new ArrayList
                {
                    new { Product = "Apple",  Price = 1.20 },
                    new { Product = "Banana", Price = 0.80 },
                    new { Product = "Orange", Price = 1.50 }
                };

                // 5. Create a WorkbookDesigner, assign the workbook and the data source
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Data", data);

                // Note: The SetRange method is not available in the current Aspose.Cells version.
                // If range‑limited processing is required, it can be achieved via other APIs.
                // For this demo we process the entire sheet.

                // 6. Process the smart markers
                designer.Process();

                // 7. Save the resulting workbook
                string outputPath = "SmartMarkerRangeResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}