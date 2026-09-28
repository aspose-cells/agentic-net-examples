// Title: How to save an Excel workbook as XLSB while retaining external data connections using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, updates a cell, sets SaveData = true on every ExternalConnection, and saves the workbook as an .xlsb using XlsbSaveOptions. | Show how to configure XlsbSaveOptions (including ExportAllColumnIndexes) to export a workbook to XLSB format without losing external connection settings.
// Common Searches: asp.net load workbook modify cell keep external connections when saving as xlsb | Aspose.Cells C# preserve data connections during XLSB export | set SaveData property for external connections before using XlsbSaveOptions | export Excel file to binary format with external data connections using Aspose.Cells | how to use XlsbSaveOptions to retain external connections in C#
// Tags: Aspose.Cells XLSB export with external connections | C# XlsbSaveOptions ExportAllColumnIndexes | ExternalConnection SaveData property Aspose.Cells | modify worksheet cell before XLSB save | preserve data connections on workbook export

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads an existing .xlsx workbook, updates cell A1, enables SaveData on each external connection, and saves the workbook as an .xlsb file using XlsbSaveOptions with ExportAllColumnIndexes set to true.
class SaveWorkbookAsXlsb
{
    static void Main()
    {
        // Load an existing workbook that contains external connections.
        // Replace "input.xlsx" with the actual source file path.
        Workbook workbook = new Workbook("input.xlsx");

        // Perform any required modifications.
        // Example: write a value to cell A1 of the first worksheet.
        workbook.Worksheets[0].Cells["A1"].PutValue("Modified");

        // Preserve external connection settings by ensuring the data fetched
        // over each connection is saved with the workbook.
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            connection.SaveData = true;
        }

        // Create XLSB save options. The default ExportAllColumnIndexes is true,
        // but it is set explicitly for clarity.
        XlsbSaveOptions saveOptions = new XlsbSaveOptions
        {
            ExportAllColumnIndexes = true
        };

        // Save the workbook as an XLSB file while retaining all external
        // connection configurations.
        workbook.Save("output.xlsb", saveOptions);
    }
}
