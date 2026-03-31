using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class ManageExternalConnections
{
    static void Main()
    {
        const string mainPath = "MainWorkbook.xlsx";
        const string externalPath = "ExternalData.xlsb";
        const string resultPath = "ResultWorkbook.xlsx";

        // Ensure the main workbook exists; create a simple one if it doesn't.
        if (!File.Exists(mainPath))
        {
            var wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Sample");
            wb.Save(mainPath, SaveFormat.Xlsx);
        }

        // Ensure the external workbook exists; create a simple one if it doesn't.
        if (!File.Exists(externalPath))
        {
            var wbExt = new Workbook();
            wbExt.Worksheets[0].Cells["A1"].PutValue(123);
            wbExt.Save(externalPath, SaveFormat.Xlsb);
        }

        // Load the primary workbook (XLSX)
        Workbook mainWorkbook = new Workbook(mainPath);

        // Load the external workbook (XLSB)
        Workbook externalWorkbook = new Workbook(externalPath);

        // Iterate through external connections in the main workbook
        // and set the OnlyUseConnectionFile flag if needed
        foreach (ExternalConnection connection in mainWorkbook.DataConnections)
        {
            connection.OnlyUseConnectionFile = true;
        }

        // Update the linked data source using the external workbook
        mainWorkbook.UpdateLinkedDataSource(new Workbook[] { externalWorkbook });

        // Recalculate formulas to reflect the refreshed data
        mainWorkbook.CalculateFormula();

        // Save the modified workbook back to XLSX format
        mainWorkbook.Save(resultPath, SaveFormat.Xlsx);
    }
}