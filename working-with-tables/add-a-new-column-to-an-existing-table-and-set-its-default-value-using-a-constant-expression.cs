// Title: Add a Column to an Aspose.Cells ListObject and Apply a Constant Default Value (C#)
// Description: Demonstrates how to insert a new column into an existing ListObject, expand the table range automatically, name the column, and set a constant default value using SetCustomCalculatedFormula. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells | C# | ListObject add column | SetCustomCalculatedFormula | InsertColumn | constant default value | expand table range | Excel automation | worksheet table column | global
// Common Searches: how to add a column to a ListObject in Aspose.Cells | set constant default value for a table column using Aspose.Cells C# | Aspose.Cells InsertColumn expand table automatically | SetCustomCalculatedFormula example Aspose.Cells | add column with default numeric value in Excel via Aspose
// Developer Intent: Insert a new column into an existing Aspose.Cells table and populate every row with the same constant value using a formula.
// Use Cases: Add a "Status" column to an orders table and default every row to "Pending". | Create a "TaxRate" column in a pricing sheet and fill all rows with the constant 0.05. | Insert a "ProcessedFlag" column into a log table and set the default value to 0 for all entries.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a new column into a ListObject and sets a constant default value of 100 for all rows using SetCustomCalculatedFormula. | Show how to add a column named "Category" to an existing Aspose.Cells table and apply the constant string "General" as the default for each data row. | Provide an example that expands a table range after inserting a column and assigns a numeric constant default using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to insert a new column into an existing ListObject, expand the table range automatically, name the column, and set a constant default value using SetCustomCalculatedFormula. The workbook is then saved as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some initial data that will become a table (2 columns)
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(20);

        // Create a table that covers the range A1:B3
        int tableIdx = sheet.ListObjects.Add("A1", "B3", true);
        ListObject table = sheet.ListObjects[tableIdx];

        // Insert a new column at index 1 (between the existing columns)
        // The 'true' flag updates references so the table expands to include the new column
        sheet.Cells.InsertColumn(1, true);

        // Access the newly added column inside the table (now the second column)
        ListColumn newCol = table.ListColumns[1];
        newCol.Name = "DefaultValue";

        // Set a constant expression as the default value for the column.
        // The formula "=5" will place the constant value 5 in every data row of this column.
        newCol.SetCustomCalculatedFormula("=5", false, false);

        // Save the workbook
        workbook.Save("AddColumnWithDefault.xlsx");
    }
}
