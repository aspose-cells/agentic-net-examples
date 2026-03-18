using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.Pivot;

class ExternalDataConnectionsDemo
{
    static void Main()
    {
        // 1. Reference an external workbook directly in a formula.
        UseExternalWorkbookLink();

        // 2. Add an external link to another workbook via the ExternalLinks collection.
        UseExternalLinksCollection();

        // 3. Update external data sources referenced by formulas.
        UpdateLinkedDataSources();

        // 4. Calculate formulas that use INDIRECT with external named ranges.
        UseCalculationOptionsWithLinkedDataSources();

        // 5. Work with external connections in a PivotTable.
        UsePivotTableExternalConnection();
    }

    // Demonstrates a formula that pulls data from an external workbook.
    static void UseExternalWorkbookLink()
    {
        Workbook wb = new Workbook();                                   // create workbook
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].Formula = "=[ExternalData.xlsx]Sheet1!$B$2";     // external reference
        wb.Save("ExternalWorkbookLink.xlsx");                           // save workbook
    }

    // Shows how to add and modify an external link using the ExternalLinks collection.
    static void UseExternalLinksCollection()
    {
        Workbook wb = new Workbook();                                   // create workbook
        // Add a link to "DataSource.xlsx" referencing Sheet1 and Sheet2
        int index = wb.Worksheets.ExternalLinks.Add("DataSource.xlsx", new string[] { "Sheet1", "Sheet2" });
        // Optionally change the data source path after adding
        wb.Worksheets.ExternalLinks[index].DataSource = "UpdatedDataSource.xlsx";
        wb.Save("ExternalLinksCollectionDemo.xlsx");                    // save workbook
    }

    // Updates the values of external links by providing the source workbooks.
    static void UpdateLinkedDataSources()
    {
        // Main workbook that contains an external reference.
        Workbook mainWb = new Workbook();
        Worksheet mainWs = mainWb.Worksheets[0];
        mainWs.Cells["A1"].Formula = "=[ExternalSource.xlsx]Sheet1!A1";

        // External workbook that supplies the data.
        Workbook externalWb = new Workbook();
        externalWb.Worksheets[0].Cells["A1"].Value = "New Value";

        // Refresh the external link using the external workbook.
        mainWb.UpdateLinkedDataSource(new Workbook[] { externalWb });

        // Recalculate to apply the refreshed value.
        mainWb.CalculateFormula();

        mainWb.Save("UpdatedLinkedDataSourceDemo.xlsx");               // save result
    }

    // Uses CalculationOptions.LinkedDataSources to resolve INDIRECT formulas that point to external workbooks.
    static void UseCalculationOptionsWithLinkedDataSources()
    {
        // External workbook with a named range.
        Workbook extWb = new Workbook();
        extWb.Worksheets[0].Cells["A2"].PutValue(123);
        int nameIdx = extWb.Worksheets.Names.Add("ExtRange");
        extWb.Worksheets.Names[nameIdx].RefersTo = "=Sheet1!$A$2";

        // Main workbook that references the external named range via INDIRECT.
        Workbook mainWb = new Workbook();
        Worksheet mainWs = mainWb.Worksheets[0];
        mainWs.Cells["A1"].Formula = "=INDIRECT(\"[ExternalWorkbook.xlsx]ExtRange\")";

        // Link the external workbook so the formula can be evaluated.
        CalculationOptions options = new CalculationOptions();
        options.LinkedDataSources = new Workbook[] { extWb };

        // Perform calculation with the linked data source.
        mainWb.CalculateFormula(options);

        mainWb.Save("CalculationOptionsLinkedDataSourcesDemo.xlsx");   // save workbook
    }

    // Retrieves and manipulates external connections associated with a PivotTable.
    static void UsePivotTableExternalConnection()
    {
        // Create a workbook with sample data.
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Amount");
        ws.Cells["A2"].PutValue("Food");
        ws.Cells["B2"].PutValue(100);
        ws.Cells["A3"].PutValue("Travel");
        ws.Cells["B3"].PutValue(200);

        // Add a PivotTable based on the sample data.
        int pivotIdx = ws.PivotTables.Add("A1:B3", "D1", "PivotTable1");
        PivotTable pt = ws.PivotTables[pivotIdx];

        // Obtain any external connections used by the PivotTable.
        ExternalConnection[] connections = pt.GetSourceDataConnections();
        if (connections != null && connections.Length > 0)
        {
            ExternalConnection conn = connections[0];
            // Example modification: enable background refresh.
            conn.BackgroundRefresh = true;
        }

        wb.Save("PivotTableExternalConnectionDemo.xlsx");              // save workbook
    }
}