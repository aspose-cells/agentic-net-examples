using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerFormulaParameterDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a template workbook with smart markers.
            Workbook template = new Workbook();
            Worksheet ws = template.Worksheets[0];
            Cells cells = ws.Cells;

            // Smart marker that will be replaced with a data value.
            cells["A2"].PutValue("&=$Data.Value");

            // Formula that uses the smart marker as a parameter.
            // Use PutValue to avoid formula parsing before processing.
            cells["B2"].PutValue("=A2*(&=$Data.Multiplier)");

            // Define the range that contains smart markers (required for range‑based processing).
            cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

            // Save the template to a memory stream.
            using (MemoryStream templateStream = new MemoryStream())
            {
                template.Save(templateStream, SaveFormat.Xlsx);
                templateStream.Position = 0;

                // 2. Load the workbook with formula parsing enabled.
                LoadOptions loadOptions = new LoadOptions
                {
                    ParsingFormulaOnOpen = true
                };
                Workbook workbook = new Workbook(templateStream, loadOptions);

                // 3. Prepare a data source for the smart markers.
                DataTable dataTable = new DataTable("Data");
                dataTable.Columns.Add("Value", typeof(double));
                dataTable.Columns.Add("Multiplier", typeof(double));
                dataTable.Rows.Add(10.0, 2.0);
                dataTable.Rows.Add(20.0, 3.0);

                // 4. Configure WorkbookDesigner to process smart markers.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    CalculateFormula = true
                };
                designer.SetDataSource(dataTable);
                designer.Process();

                // 5. Save the processed workbook.
                workbook.Save("Result.xlsx");
            }
        }
    }
}