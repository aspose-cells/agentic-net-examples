using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerRangeExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set up header row
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Value");

                // Insert smart markers for the data rows (starting from row 2)
                // The range will later be limited to the first 100 rows
                sheet.Cells["A2"].PutValue("&=Data.Name");
                sheet.Cells["B2"].PutValue("&=Data.Value");

                // Prepare a data source with many rows (e.g., 500 rows)
                DataTable data = new DataTable("Data");
                data.Columns.Add("Name", typeof(string));
                data.Columns.Add("Value", typeof(double));
                for (int i = 1; i <= 500; i++)
                {
                    data.Rows.Add($"Item {i}", i * 10.0);
                }

                // Initialize WorkbookDesigner and assign the workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Set the data source for the smart markers
                designer.SetDataSource(data);

                // Define a range that covers only the first 100 data rows (A2:B101)
                // Naming the range with the special marker name enables processing of only this range
                Aspose.Cells.Range smartMarkerRange = sheet.Cells.CreateRange("A2:B101");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // Process only the defined range; true = preserve unrecognized markers (not needed here)
                designer.Process(smartMarkerRange, true);

                // Save the resulting workbook
                workbook.Save("SmartMarkersFirst100Rows.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}