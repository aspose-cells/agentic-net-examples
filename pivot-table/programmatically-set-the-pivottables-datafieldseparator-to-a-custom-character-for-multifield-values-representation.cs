// Title: Aspose.Cells C# – Set a Custom DataFieldSeparator for a PivotTable
// Description: Demonstrates how to assign a custom character to the PivotTable.DataFieldSeparator property in Aspose.Cells before refreshing the table, enabling multi‑field values to be displayed with any delimiter and saving the workbook as an Excel file.
// Keywords: Aspose.Cells PivotTable DataFieldSeparator | C# custom delimiter pivot table | Aspose.Cells multi‑field pivot separator | DataFieldSeparator property example | Change pivot data field separator .NET | Custom character for pivot values Aspose | Excel pivot table custom separator C# | Aspose.Cells API DataFieldSeparator | PivotTable custom delimiter code | Aspose.Cells GitHub example DataFieldSeparator
// Common Searches: Aspose.Cells set DataFieldSeparator C# | PivotTable custom delimiter Aspose.Cells | How to change data field separator in Aspose pivot | C# example for multi‑field pivot separator | Aspose.Cells DataFieldSeparator usage | Set pipe character as pivot data separator Aspose | Aspose.Cells pivot table custom separator tutorial
// Developer Intent: Configure a custom character for the PivotTable.DataFieldSeparator property in C# to control how combined data fields appear.
// Use Cases: Show sales and quantity side‑by‑side in a single cell using a pipe (|) separator. | Export pivot results to CSV where a semicolon (;) delimiter simplifies downstream parsing. | Create readable Excel reports by inserting a newline character as the separator for multi‑field values.
// AI Prompts: Provide C# code that sets PivotTable.DataFieldSeparator to a custom character before calling RefreshData in Aspose.Cells. | Show an Aspose.Cells example adding two data fields to a PivotTable and using ';' as the separator so cells display "Sum1;Sum2". | Explain how to read, modify, and persist the DataFieldSeparator property of an existing PivotTable in a workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Demonstrates how to assign a custom character to the PivotTable.DataFieldSeparator property in Aspose.Cells before refreshing the table, enabling multi‑field values to be displayed with any delimiter and saving the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].Value = "Category";
                cells["B1"].Value = "SubCategory";
                cells["C1"].Value = "Amount";

                cells["A2"].Value = "Fruit";
                cells["B2"].Value = "Apple";
                cells["C2"].Value = 120;

                cells["A3"].Value = "Fruit";
                cells["B3"].Value = "Banana";
                cells["C3"].Value = 80;

                cells["A4"].Value = "Vegetable";
                cells["B4"].Value = "Carrot";
                cells["C4"].Value = 50;

                cells["A5"].Value = "Vegetable";
                cells["B5"].Value = "Broccoli";
                cells["C5"].Value = 70;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "MyPivotTable");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh the pivot table data (using the available API)
                pivotTable.RefreshData();

                // Calculate the pivot table data
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTableWithCustomSeparator.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
