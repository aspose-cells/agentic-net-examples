using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsJsonMergeDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and add a smart marker that will consume the merged JSON data.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            // Smart marker expects a collection named "Employees" with fields "Name" and "Age".
            sheet.Cells["A1"].PutValue("&=$Employees.Name");
            sheet.Cells["B1"].PutValue("&=$Employees.Age");

            // 2. Prepare multiple JSON arrays that need to be merged.
            string jsonArray1 = @"[
                { ""Name"": ""John Doe"", ""Age"": 30 },
                { ""Name"": ""Jane Smith"", ""Age"": 25 }
            ]";

            string jsonArray2 = @"[
                { ""Name"": ""Mike Johnson"", ""Age"": 40 },
                { ""Name"": ""Emily Davis"", ""Age"": 35 }
            ]";

            // 3. Merge the arrays into a single JSON array.
            // Remove the surrounding brackets and concatenate the inner objects with commas.
            string mergedJson = "[" +
                jsonArray1.Trim().TrimStart('[').TrimEnd(']') + "," +
                jsonArray2.Trim().TrimStart('[').TrimEnd(']') +
                "]";

            // 4. Set up the WorkbookDesigner, assign the merged JSON as a data source,
            //    and process the smart markers.
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            // The data source name must match the smart marker prefix ("Employees").
            designer.SetJsonDataSource("Employees", mergedJson);
            designer.Process();

            // 5. Save the result.
            workbook.Save("MergedJsonSmartMarkers.xlsx");
        }
    }
}