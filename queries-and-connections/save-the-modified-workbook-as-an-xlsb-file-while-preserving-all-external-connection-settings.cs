using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class SaveWorkbookAsXlsb
{
    static void Main()
    {
        // Load an existing workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Preserve all external connection settings by ensuring the data is saved with the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            connection.SaveData = true;
        }

        // Configure XLSB save options as needed
        XlsbSaveOptions saveOptions = new XlsbSaveOptions
        {
            ExportAllColumnIndexes = true,
            RefreshChartCache = true,
            SortNames = true,
            SortExternalNames = true
        };

        // Save the workbook as an XLSB file while keeping external connection settings intact
        workbook.Save("output.xlsb", saveOptions);
    }
}