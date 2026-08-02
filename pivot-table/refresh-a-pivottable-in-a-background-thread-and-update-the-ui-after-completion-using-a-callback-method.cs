using System;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    class Program
    {
        // Callback method that will be invoked on the UI thread after refresh completes
        static void OnRefreshCompleted()
        {
            Console.WriteLine("Pivot table refresh completed and workbook saved.");
            // Here you could update UI controls, e.g., enable buttons or refresh views
        }

        // Asynchronous method that refreshes all pivot tables in the workbook
        static void RefreshPivotTablesAsync(string workbookPath, Action callback)
        {
            // Run the refresh operation on a background thread
            Task.Run(() =>
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(workbookPath);

                // Refresh all pivot tables in the workbook
                // This uses the provided RefreshPivotTables method from the API
                workbook.Worksheets.RefreshPivotTables();

                // Optionally calculate data after refresh (if needed)
                // Iterate through worksheets and calculate each pivot table
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pt in sheet.PivotTables)
                    {
                        pt.CalculateData();
                    }
                }

                // Save the workbook back to the same file (or a new file)
                workbook.Save(workbookPath);
            })
            .ContinueWith(t =>
            {
                // Invoke the callback on the thread that called RefreshPivotTablesAsync
                // In a real UI application you would marshal this to the UI thread (e.g., using Dispatcher)
                callback?.Invoke();
            });
        }

        static void Main(string[] args)
        {
            // Path to the Excel file containing the pivot table
            string filePath = "PivotTableDemo.xlsx";

            // Ensure the workbook exists; create a simple one if not present
            if (!System.IO.File.Exists(filePath))
            {
                // Create a workbook with sample data and a pivot table (using provided creation pattern)
                Workbook wb = new Workbook();
                Worksheet dataSheet = wb.Worksheets[0];
                dataSheet.Cells["A1"].PutValue("Product");
                dataSheet.Cells["B1"].PutValue("Sales");
                dataSheet.Cells["A2"].PutValue("Apple");
                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["A3"].PutValue("Banana");
                dataSheet.Cells["B3"].PutValue(150);
                dataSheet.Cells["A4"].PutValue("Apple");
                dataSheet.Cells["B4"].PutValue(80);

                // Add a pivot table
                int pivotIndex = dataSheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
                PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales

                // Save the initial workbook
                wb.Save(filePath);
            }

            // Start the asynchronous refresh operation
            RefreshPivotTablesAsync(filePath, OnRefreshCompleted);

            // Keep the console alive to observe the callback (in a real UI app this wouldn't be needed)
            Console.WriteLine("Refresh started on background thread. Press any key to exit...");
            Console.ReadKey();
        }
    }
}