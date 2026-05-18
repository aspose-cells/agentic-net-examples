using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            RefreshAllExternalDataConnections.Run();
        }
    }

    public class RefreshAllExternalDataConnections
    {
        public static void Run()
        {
            // Load the workbook that contains external data connections
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Get the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Ensure each connection is set to refresh on load
            foreach (ExternalConnection conn in connections)
            {
                conn.RefreshOnLoad = true;
            }

            // Refresh all external data connections, pivot tables, and charts
            workbook.RefreshAll();

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}