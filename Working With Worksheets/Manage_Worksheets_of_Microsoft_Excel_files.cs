using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetManagement
{
    class Program
    {
        static void Main(string[] args)
        {
            // -------------------------------------------------
            // 1. Create a new workbook (lifecycle: create)
            // -------------------------------------------------
            Workbook workbook = new Workbook(); // default Xlsx format

            // -------------------------------------------------
            // 2. Access the default worksheet and set a name
            // -------------------------------------------------
            Worksheet defaultSheet = workbook.Worksheets[0];
            defaultSheet.Name = "Summary";

            // -------------------------------------------------
            // 3. Add a new worksheet to the collection
            // -------------------------------------------------
            int newSheetIndex = workbook.Worksheets.Add(); // adds a blank worksheet
            Worksheet dataSheet = workbook.Worksheets[newSheetIndex];
            dataSheet.Name = "Data";

            // -------------------------------------------------
            // 4. Populate some data in the new worksheet
            // -------------------------------------------------
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Quantity");
            dataSheet.Cells["C1"].PutValue("Price");

            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["C2"].PutValue(0.5);

            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["C3"].PutValue(0.3);

            // -------------------------------------------------
            // 5. Insert a column between B and C (lifecycle: modify)
            // -------------------------------------------------
            dataSheet.Cells.InsertColumn(2); // inserts at index 2 (C becomes D)

            // Add a header for the new column
            dataSheet.Cells["C1"].PutValue("Total");
            // Fill the Total column with a formula
            dataSheet.Cells["C2"].Formula = "=B2*D2";
            dataSheet.Cells["C3"].Formula = "=B3*D3";

            // -------------------------------------------------
            // 6. Auto-fit columns for better appearance
            // -------------------------------------------------
            dataSheet.AutoFitColumns();

            // -------------------------------------------------
            // 7. Copy the "Data" worksheet to create a backup sheet
            // -------------------------------------------------
            int copyIndex = workbook.Worksheets.AddCopy("Data");
            Worksheet backupSheet = workbook.Worksheets[copyIndex];
            backupSheet.Name = "Data_Backup";

            // -------------------------------------------------
            // 8. Delete the original "Summary" worksheet (if not needed)
            // -------------------------------------------------
            // Note: Worksheet index may have changed after adding new sheets.
            // Find the sheet by name before removal.
            int summaryIndex = workbook.Worksheets.IndexOf(defaultSheet);
            if (summaryIndex >= 0)
            {
                workbook.Worksheets.RemoveAt(summaryIndex);
            }

            // -------------------------------------------------
            // 9. Set column width and row height explicitly
            // -------------------------------------------------
            // Set column A width to 20 characters
            dataSheet.Cells.SetColumnWidth(0, 20);
            // Set row 1 height to 25 points
            dataSheet.Cells.SetRowHeight(0, 25);

            // -------------------------------------------------
            // 10. Save the workbook to disk (lifecycle: save)
            // -------------------------------------------------
            string outputPath = "ManagedWorksheets.xlsx";
            workbook.Save(outputPath); // uses default SaveFormat based on file extension

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}