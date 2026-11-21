// Title: How to add a calculated item to a pivot field in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add calculated item to pivot table field example | Create custom calculated item in Excel pivot using Aspose.Cells .NET | Add calculated item to row field of pivot table with Aspose.Cells API | Programmatically insert calculated item into Excel pivot table using C# | Aspose.Cells AddCalculatedItem method usage for pivot tables
// Tags: Aspose.Cells AddCalculatedItem API | C# pivot table calculated item | Excel workbook pivot field modification | programmatic pivot table item creation | Aspose.Cells pivot table custom formula

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, finds a specific pivot table and its row field, then calls the field's AddCalculatedItem method with a custom expression to create a new calculated item, and finally saves the updated workbook.
class AddCalculatedItemExample
{
    static void Main()
    {
        // Load the existing workbook
        // (Replace with the actual path to your Excel file)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (or specify the appropriate one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the first pivot table on the worksheet (or locate by name)
        // (Replace "PivotTable1" with the actual name if needed)
        PivotTable pivotTable = null;
        foreach (PivotTable pt in worksheet.PivotTables)
        {
            if (pt.Name == "PivotTable1")
            {
                pivotTable = pt;
                break;
            }
        }

        if (pivotTable == null)
        {
            Console.WriteLine("Pivot table not found.");
            return;
        }

        // Locate the field to which the calculated item will be added
        // (Replace "FieldName" with the actual field name)
        PivotField field = null;
        foreach (PivotField pf in pivotTable.RowFields) // adjust collection if field is in column/ data area
        {
            if (pf.Name == "FieldName")
            {
                field = pf;
                break;
            }
        }

        if (field == null)
        {
            Console.WriteLine("Field not found in the pivot table.");
            return;
        }

        // Add a calculated item to the field
        // (Replace "NewItem" with the desired calculated item name
        //  and "Expression" with the formula, e.g., "[Sales] * 0.1")
        string calculatedItemName = "NewItem";
        string expression = "[Sales] * 0.1";

        // The AddCalculatedItem method creates the item with the given expression
        field.AddCalculatedItem(calculatedItemName, expression);

        // Save the workbook with the new calculated item
        // (Replace with the desired output path)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);

        Console.WriteLine("Calculated item added and workbook saved successfully.");
    }
}
