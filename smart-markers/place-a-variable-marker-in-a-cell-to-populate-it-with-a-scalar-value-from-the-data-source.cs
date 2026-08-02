using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsVariableMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Place a variable marker (smart marker) in cell A2.
            // The marker &=$ProductName will be replaced with the value from the data source.
            cells["A2"].PutValue("&=$ProductName");

            // Prepare a simple data source: a dictionary with a scalar value.
            var dataSource = new Dictionary<string, object>
            {
                { "ProductName", "Aspose.Cells Sample Product" }
            };

            // Initialize WorkbookDesigner, assign the workbook and set the data source.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Data", dataSource);

            // Process the smart markers to populate the cell with the scalar value.
            designer.Process();

            // Save the resulting workbook.
            workbook.Save("VariableMarkerOutput.xlsx");
        }
    }
}