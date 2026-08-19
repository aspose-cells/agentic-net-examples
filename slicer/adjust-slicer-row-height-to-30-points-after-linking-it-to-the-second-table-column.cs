// Title: Aspose.Cells C# – Set slicer row height to 30 pt after linking to a table column
// Description: Creates a workbook, adds a two‑column table, inserts a slicer linked to the table's second column (Category), sets the slicer.RowHeight to 30 points, and saves the file as SlicerRowHeight30.xlsx.
// Keywords: Aspose.Cells slicer row height | C# set slicer item height | RowHeight property Aspose.Cells | link slicer to table column .NET | Excel slicer formatting programmatically | Aspose.Cells example C# | adjust slicer height points | automate slicer appearance | Aspose.Cells dashboard slicer
// Common Searches: Aspose.Cells change slicer row height C# | Set slicer item height to 30 points in .NET | Link slicer to second column of a table using Aspose.Cells | RowHeight property for Excel slicer Aspose | Programmatic slicer formatting Aspose.Cells
// Developer Intent: Set the slicer’s row height to 30 points after it is linked to the second column of a table.
// Use Cases: Generate an Excel report with a slicer for the Category column where each item has a uniform 30‑pt height for readability. | Automate dashboard styling by linking a slicer to a specific table column and applying consistent row height. | Create a reusable C# routine that adds a slicer to a worksheet and standardizes its item height across multiple workbooks.
// AI Prompts: Write C# code with Aspose.Cells that adds a slicer linked to the second column of a table and sets its RowHeight to 30 points. | Show how to adjust the RowHeight property of an Aspose.Cells slicer after creation and save the workbook. | Explain the steps to link a slicer to a table column and change the slicer item height programmatically in .NET.

using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

// Creates a workbook, adds a two‑column table, inserts a slicer linked to the table's second column (Category), sets the slicer.RowHeight to 30 points, and saves the file as SlicerRowHeight30.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a table with two columns
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue("Fruit");
        worksheet.Cells["A3"].PutValue("Carrot");
        worksheet.Cells["B3"].PutValue("Vegetable");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B4"].PutValue("Fruit");

        // Add a table (ListObject) covering the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Add a slicer linked to the second column (Category) of the table
        // Position the slicer at cell A6 (row index 5, column index 0)
        SlicerCollection slicers = worksheet.Slicers;
        int slicerIndex = slicers.Add(table, table.ListColumns[1], 5, 0);
        Slicer slicer = slicers[slicerIndex];

        // Adjust the row height of each slicer item to 30 points
        slicer.RowHeight = 30;

        // Save the workbook
        workbook.Save("SlicerRowHeight30.xlsx");
    }
}
