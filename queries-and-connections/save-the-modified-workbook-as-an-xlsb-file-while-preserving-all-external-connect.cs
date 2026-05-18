using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class SaveWorkbookAsXlsb
{
    static void Main()
    {
        // Load an existing workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Example modification: write a value to a cell
        workbook.Worksheets[0].Cells["A1"].PutValue("Modified");

        // Preserve all external connection settings by ensuring the data is saved with the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            connection.SaveData = true;
        }

        // Configure XLSB save options
        XlsbSaveOptions saveOptions = new XlsbSaveOptions
        {
            ExportAllColumnIndexes = true,   // export all column indexes
            SortExternalNames = true,        // keep external names sorted
            RefreshChartCache = true         // refresh chart cache if any
        };

        // Save the workbook as an XLSB file while preserving external connections
        workbook.Save("output.xlsb", saveOptions);
    }
}