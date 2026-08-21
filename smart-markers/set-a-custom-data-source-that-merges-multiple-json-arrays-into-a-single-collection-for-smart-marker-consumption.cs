// Title: Merge Multiple JSON Arrays into a Single Smart Marker Data Source with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, place smart markers, merge two JSON array strings into one collection, set the combined JSON as a data source via WorkbookDesigner.SetJsonDataSource, process the markers, and save the resulting Excel file.
// Keywords: Aspose.Cells | C# | .NET | Smart markers | JSON merge | WorkbookDesigner | SetJsonDataSource | combined JSON data source | Excel automation
// Common Searches: Aspose.Cells merge JSON arrays | Set JSON data source for smart markers .NET | WorkbookDesigner SetJsonDataSource example | Combine multiple JSON collections in Aspose.Cells | Smart marker data source from merged JSON
// Developer Intent: Combine several JSON arrays into a single data source that can be consumed by Aspose.Cells smart markers in a C# application.
// Use Cases: Unify employee and contractor JSON lists into one "People" collection for a single smart‑marker table in a HR report. | Aggregate product catalog fragments before populating an invoice worksheet with smart markers. | Consolidate survey response arrays into a single data source to generate summary charts via smart markers.
// AI Prompts: Generate C# code that merges three JSON arrays and assigns the result to a smart‑marker data source using Aspose.Cells WorkbookDesigner. | Explain step‑by‑step how WorkbookDesigner.SetJsonDataSource works with a merged JSON string and how to process smart markers in a .NET project. | Provide a robust method to handle empty or null JSON arrays when merging them for a smart‑marker data source in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsJsonMergeExample
{
    // Shows how to create a workbook, place smart markers, merge two JSON array strings into one collection, set the combined JSON as a data source via WorkbookDesigner.SetJsonDataSource, process the markers, and save the resulting Excel file.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Place smart markers that will consume the merged JSON collection
            //    The marker syntax "&=$DataSource.ColumnName"
            sheet.Cells["A1"].PutValue("&=$People.Name");
            sheet.Cells["B1"].PutValue("&=$People.Age");

            // 4. Prepare two separate JSON arrays
            string jsonArray1 = "[{\"Name\":\"John\",\"Age\":30}]";
            string jsonArray2 = "[{\"Name\":\"Jane\",\"Age\":25}]";

            // 5. Merge the arrays into a single JSON array string
            //    Remove the surrounding brackets and concatenate with a comma
            string mergedJson = "[" +
                jsonArray1.Trim('[', ']') + "," +
                jsonArray2.Trim('[', ']') +
                "]";

            // 6. Create a WorkbookDesigner and assign the workbook (lifecycle load)
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // 7. Set the merged JSON as a data source named "People"
            //    This will be used by the smart markers defined earlier
            designer.SetJsonDataSource("People", mergedJson);

            // 8. Process the smart markers to populate data
            designer.Process();

            // 9. Save the resulting workbook (lifecycle save)
            workbook.Save("MergedJsonSmartMarkers.xlsx");
        }
    }
}
