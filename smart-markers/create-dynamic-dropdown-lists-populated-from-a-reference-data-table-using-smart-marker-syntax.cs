// Title: Create a dynamic Excel dropdown list with Aspose.Cells smart markers referencing a hidden worksheet in C#
// Common Searches: how to use Aspose.Cells smart markers to create a dropdown list from a hidden worksheet in C# | Aspose.Cells ?= syntax for data validation list referencing another sheet | generate Excel dropdown list dynamically with reference data using Aspose.Cells C# | C# example of creating a hidden reference sheet for Excel data validation with Aspose.Cells | smart marker range reference for Excel dropdown using Aspose.Cells WorkbookDesigner
// Tags: smart marker data validation range | Aspose.Cells hidden reference sheet | WorkbookDesigner process smart markers C# | Excel dropdown list from range using Aspose.Cells | ?= syntax data validation Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, adds a hidden sheet named "ReferenceData" populated with fruit items, inserts a smart marker "?=ReferenceData!A1:A5" into cell B2 of the main sheet to define a data‑validation dropdown, processes the marker with WorkbookDesigner, and saves the file as DynamicDropdown.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // 1. Create a hidden sheet that holds reference data
        // -------------------------------------------------
        Worksheet refSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        refSheet.Name = "ReferenceData";

        // Sample list items for the dropdown
        string[] listItems = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
        for (int i = 0; i < listItems.Length; i++)
        {
            // Populate column A with the items
            refSheet.Cells[i, 0].PutValue(listItems[i]);
        }

        // -------------------------------------------------
        // 2. Create the main sheet where the dropdown will appear
        // -------------------------------------------------
        Worksheet mainSheet = workbook.Worksheets[0];
        mainSheet.Name = "Main";

        // Place a smart marker in the target cell.
        // The syntax "?=ReferenceData!A1:A5" tells Aspose.Cells to create a data validation list
        // that references the range A1:A5 on the ReferenceData sheet.
        mainSheet.Cells["B2"].PutValue("?=ReferenceData!A1:A5");

        // -------------------------------------------------
        // 3. Process smart markers
        // -------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;
        // No external data source is required because the smart marker directly references a range.
        designer.Process();

        // -------------------------------------------------
        // 4. Save the workbook
        // -------------------------------------------------
        workbook.Save("DynamicDropdown.xlsx");
    }
}
