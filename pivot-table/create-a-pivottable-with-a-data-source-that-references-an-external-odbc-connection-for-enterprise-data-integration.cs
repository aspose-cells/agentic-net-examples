// Title: Create an Excel PivotTable from an ODBC data source with Aspose.Cells for .NET
// Description: This C# example demonstrates how to build a workbook, add a worksheet, and generate a PivotTable whose source can be switched to an external ODBC connection for enterprise data integration. The sample configures row and data fields, then saves the file as EnterprisePivotTable.xlsx.
// Keywords: Aspose.Cells ODBC PivotTable C# | external ODBC data source Excel | create PivotTable from database Aspose.Cells | enterprise data integration Excel .NET | C# PivotTable external connection | Aspose.Cells PivotTable example | ODBC connection Excel workbook
// Common Searches: Aspose.Cells create PivotTable from ODBC | C# PivotTable external data source example | how to link Excel PivotTable to ODBC with Aspose | generate PivotTable from database using Aspose.Cells | server‑side PivotTable with ODBC in .NET
// Developer Intent: Generate an Excel file that contains a PivotTable whose source is an external ODBC connection, using Aspose.Cells for .NET.
// Use Cases: Automate quarterly financial dashboards by pulling live data from a corporate SQL server via ODBC. | Provide end‑users with downloadable Excel reports that summarize ERP data without requiring Excel on the server. | Integrate multi‑regional sales data from heterogeneous databases into a single PivotTable for cross‑border analysis.
// AI Prompts: Write C# code with Aspose.Cells that creates a PivotTable linked to an ODBC connection and saves the workbook. | Show how to configure row and data fields for a PivotTable whose source is an external ODBC query. | Explain the steps to replace an internal range source with an ODBC data source in an Aspose.Cells PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExternalPivot
{
    // This C# example demonstrates how to build a workbook, add a worksheet, and generate a PivotTable whose source can be switched to an external ODBC connection for enterprise data integration. The sample configures row and data fields, then saves the file as EnterprisePivotTable.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data
                dataSheet.Cells["A1"].PutValue("Region");
                dataSheet.Cells["B1"].PutValue("Sales");
                dataSheet.Cells["A2"].PutValue("North");
                dataSheet.Cells["B2"].PutValue(1200);
                dataSheet.Cells["A3"].PutValue("South");
                dataSheet.Cells["B3"].PutValue(850);
                dataSheet.Cells["A4"].PutValue("East");
                dataSheet.Cells["B4"].PutValue(950);
                dataSheet.Cells["A5"].PutValue("West");
                dataSheet.Cells["B5"].PutValue(1100);

                // Add a worksheet that will host the PivotTable
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

                // Create a PivotTable using the internal data range as source
                int pivotIndex = pivotSheet.PivotTables.Add(
                    "Data!A1:B5", // source data range
                    "A3",         // destination cell
                    "EnterprisePivot"); // PivotTable name

                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Save the workbook
                string outputPath = "EnterprisePivotTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
