using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class ExternalConnectionDemo
{
    static void Main()
    {
        // Load the source XLSX workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Access the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate through each connection and toggle the OnlyUseConnectionFile property
        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];
            conn.OnlyUseConnectionFile = !conn.OnlyUseConnectionFile;
        }

        // Save the modified workbook as XLSX
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}