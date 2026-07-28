// Title: Create a Named Range from AdvancedFilter and Use It as a Slicer Source in Aspose.Cells (C#)
// Description: Step‑by‑step C# example that applies Aspose.Cells AdvancedFilter, defines a named range for the filtered list, converts it to a ListObject table, adds a slicer linked to the first column, and saves the workbook.
// Keywords: Aspose.Cells C# | .NET Excel automation | AdvancedFilter example | named range Excel | slicer data source | ListObject table | create slicer programmatically | Excel workbook sample code | GitHub Aspose.Cells example | dynamic Excel report
// Common Searches: Aspose.Cells create slicer from filtered data | C# AdvancedFilter to named range | how to add slicer to ListObject using Aspose.Cells | use named range as slicer source .NET | sample code for AdvancedFilter and slicer Aspose.Cells
// Developer Intent: Generate a named range from an AdvancedFilter result, turn it into a table, and attach a slicer to the table’s first column using Aspose.Cells for .NET.
// Use Cases: Build an interactive Excel dashboard where a slicer controls a filtered dataset. | Create reusable named ranges for filtered reports that can be referenced by multiple slicers or formulas. | Automate generation of category‑specific worksheets with a slicer for quick user navigation.
// AI Prompts: Write C# code with Aspose.Cells that runs AdvancedFilter, creates a named range for the output, converts it to a ListObject, adds a slicer on the first column, and saves the file. | Explain how to set the RefersTo property of a workbook name to a filtered range and bind that name to a slicer in Aspose.Cells. | Provide a concise tutorial for programmatically adding a slicer to a filtered table in an Excel workbook using Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // Step‑by‑step C# example that applies Aspose.Cells AdvancedFilter, defines a named range for the filtered list, converts it to a ListObject table, adds a slicer linked to the first column, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Header + rows)
            // Columns: Category | Value
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Vegetable");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Fruit");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("Grain");
            sheet.Cells["B5"].PutValue(15);
            sheet.Cells["A6"].PutValue("Fruit");
            sheet.Cells["B6"].PutValue(25);

            // -----------------------------------------------------------------
            // Define criteria range for AdvancedFilter (filter Category = "Fruit")
            // D1: Header, D2: Criteria
            sheet.Cells["D1"].PutValue("Category");
            sheet.Cells["D2"].PutValue("Fruit");

            // Apply AdvancedFilter to copy filtered rows to F1 (including headers)
            // Parameters: isFilter = false (copy), listRange = "A1:B6",
            // criteriaRange = "D1:D2", copyTo = "F1", uniqueRecordOnly = false
            sheet.AdvancedFilter(false, "A1:B6", "D1:D2", "F1", false);

            // -----------------------------------------------------------------
            // Create a named range that refers to the copied filtered list (F1:G6)
            int nameIndex = workbook.Worksheets.Names.Add("FilteredList");
            workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$F$1:$G$6";

            // -----------------------------------------------------------------
            // Convert the filtered list area into a ListObject (table)
            // Start at row 0, column 5 (cell F1), end at row 5, column 6 (cell G6)
            int listObjectIndex = sheet.ListObjects.Add(0, 5, 5, 6, true);
            ListObject listObject = sheet.ListObjects[listObjectIndex];
            listObject.DisplayName = "FilteredTable";

            // -----------------------------------------------------------------
            // Add a slicer that uses the first column of the ListObject as its source
            // Destination cell for slicer upper‑left corner is I1
            int slicerIndex = sheet.Slicers.Add(listObject, 0, "I1");
            Slicer slicer = sheet.Slicers[slicerIndex];
            slicer.Caption = "Category Slicer";

            // Save the workbook
            workbook.Save("SlicerWithFilteredNamedRange.xlsx");
        }
    }
}
