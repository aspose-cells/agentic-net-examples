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

        // Preserve external connection settings by ensuring the data is saved with the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            connection.SaveData = true;
        }

        // Create XLSB save options
        XlsbSaveOptions saveOptions = new XlsbSaveOptions
        {
            ExportAllColumnIndexes = true // default, kept for clarity
        };

        // Save the workbook as XLSB while keeping external connection settings
        workbook.Save("output.xlsb", saveOptions);
    }
}