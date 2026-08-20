// Title: C# Example: ShowRowHeaderCaption Property to Display Row Header Labels in Aspose.Cells PivotTable
// Description: Demonstrates creating a workbook, populating sample data, adding a PivotTable, assigning row, column, and data fields, enabling the ShowRowHeaderCaption property, refreshing and calculating the PivotTable, and saving the file as PivotTable_ShowRowHeaders.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# PivotTable | ShowRowHeaderCaption | row header labels | .NET Excel pivot | display row headers | Aspose.Cells example | PivotTable properties | Excel export | GitHub sample
// Common Searches: Aspose.Cells ShowRowHeaderCaption C# | how to display row headers in Aspose.Cells pivot table | C# code example for PivotTable row header captions | Aspose.Cells PivotTable properties tutorial | enable row header labels in Excel pivot using Aspose
// Developer Intent: Enable the ShowRowHeaderCaption property so the PivotTable shows row header captions.
// Use Cases: Generate financial reports with clear row headings for better readability. | Create Excel dashboards where row labels improve data comprehension. | Provide end‑users with readable pivot tables in .NET web or desktop applications. | Allow dynamic toggling of row header visibility based on user preferences.
// AI Prompts: Write C# code using Aspose.Cells to create a PivotTable and turn on ShowRowHeaderCaption. | Explain the impact of ShowRowHeaderCaption on PivotTable layout and how to disable it. | Provide a step‑by‑step guide to add a PivotTable with row header captions and customize its appearance in Aspose.Cells for .NET. | Suggest how to conditionally set ShowRowHeaderCaption based on a configuration setting.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Demonstrates creating a workbook, populating sample data, adding a PivotTable, assigning row, column, and data fields, enabling the ShowRowHeaderCaption property, refreshing and calculating the PivotTable, and saving the file as PivotTable_ShowRowHeaders.xlsx using Aspose.Cells for .NET.
    class ShowRowHeadersExample
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (data source)
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Item");
            dataSheet.Cells["C1"].PutValue("Amount");

            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(120);

            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(80);

            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(50);

            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue("Tomato");
            dataSheet.Cells["C5"].PutValue(70);

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table using the data range and place it at D3
            int pivotIndex = pivotSheet.PivotTables.Add("=Sheet1!A1:C5", "D3", "MyPivotTable");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Item");    // Column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");   // Data field

            // Enable the display of row header captions (row header labels)
            pivotTable.ShowRowHeaderCaption = true;

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_ShowRowHeaders.xlsx");
        }
    }
}
