using System;
using System.Data;
using System.IO;
using Aspose.Cells;

class SmartMarkerOverflowDemo
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Generate a large CSV data set that exceeds a single sheet's row limit.
        // ------------------------------------------------------------
        var csvBuilder = new System.Text.StringBuilder();
        csvBuilder.AppendLine("Name,Value"); // header row
        for (int i = 0; i < 2000; i++) // create many rows to force overflow
        {
            csvBuilder.AppendLine($"Item{i},{i}");
        }
        byte[] csvBytes = System.Text.Encoding.UTF8.GetBytes(csvBuilder.ToString());

        // ------------------------------------------------------------
        // 2. Load the CSV into a workbook with automatic overflow enabled.
        //    TxtLoadOptions.ExtendToNextSheet = true makes extra rows spill to new sheets.
        // ------------------------------------------------------------
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            ExtendToNextSheet = true
        };
        Workbook workbook = new Workbook(new MemoryStream(csvBytes), loadOptions);

        // ------------------------------------------------------------
        // 3. Insert smart markers into the first worksheet.
        //    The markers reference a data source named "Data".
        // ------------------------------------------------------------
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("&=Data.Name");
        ws.Cells["B1"].PutValue("&=Data.Value");

        // Define the range that contains the smart markers.
        // Naming it "_CellsSmartMarkers" tells the designer to treat it as a range marker.
        ws.Cells.CreateRange("A1:B1").Name = "_CellsSmartMarkers";

        // ------------------------------------------------------------
        // 4. Build a DataTable that matches the smart marker fields.
        // ------------------------------------------------------------
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Value", typeof(int));
        for (int i = 0; i < 2000; i++)
        {
            dt.Rows.Add($"Item{i}", i);
        }

        // ------------------------------------------------------------
        // 5. Configure WorkbookDesigner to use range smart markers (LineByLine = false)
        //    and set the data source.
        // ------------------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            LineByLine = false // use range smart markers
        };
        designer.SetDataSource(dt);

        // ------------------------------------------------------------
        // 6. Process the smart markers.
        //    Because the data exceeds the row capacity of the first sheet,
        //    Aspose.Cells automatically creates additional worksheets and continues
        //    populating the data there.
        // ------------------------------------------------------------
        designer.Process();

        // ------------------------------------------------------------
        // 7. Save the resulting workbook.
        // ------------------------------------------------------------
        workbook.Save("SmartMarkerOverflowResult.xlsx");
    }
}