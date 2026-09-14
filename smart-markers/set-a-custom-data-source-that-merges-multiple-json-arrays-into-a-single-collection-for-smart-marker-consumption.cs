// Title: Merge multiple JSON arrays into one collection and use it as a smart‑marker data source with Aspose.Cells in C#
// AI Prompts: Combine two JSON array strings into a single JSON array and assign it to WorkbookDesigner.SetJsonDataSource for the "People" smart marker. | Create an Excel workbook, place a smart marker that references &=$People.Name, set the merged JSON as the data source, process the markers, and save the workbook. | Show how to handle potential errors while merging JSON arrays and populating smart markers using Aspose.Cells for .NET.
// Common Searches: c# aspocells merge json arrays for smart marker data source | set json data source for smart markers using WorkbookDesigner in .NET | combine multiple JSON collections into one for Aspose.Cells smart markers | populate smart markers from merged JSON in C#
// Tags: merge json arrays with Aspose.Cells WorkbookDesigner | setjsondatasource merged collection C# | smart marker json data source merging | excel generation from combined json using Aspose.Cells | c# smart markers multiple json arrays

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook with a smart marker that iterates over a "People" collection, merges two JSON arrays into a single JSON array string, sets this merged JSON as the data source via WorkbookDesigner.SetJsonDataSource, processes the smart markers, and saves the populated Excel file.
    public class MergeJsonArraysForSmartMarkers
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // 1. Create a new workbook (template) and add a smart marker.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            // Smart marker expects a collection named "People" with a "Name" field.
            sheet.Cells["A1"].PutValue("&=$People.Name");

            // 2. Prepare two separate JSON arrays.
            string jsonArray1 = "[{\"Name\":\"John\"},{\"Name\":\"Jane\"}]";
            string jsonArray2 = "[{\"Name\":\"Bob\"},{\"Name\":\"Alice\"}]";

            // 3. Merge the arrays into a single JSON array.
            // Remove the surrounding brackets, concatenate with a comma, and wrap again.
            string mergedJson = "[" +
                jsonArray1.TrimStart('[').TrimEnd(']') + "," +
                jsonArray2.TrimStart('[').TrimEnd(']') +
                "]";

            // 4. Set the merged JSON as a data source for the smart marker.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetJsonDataSource("People", mergedJson);

            // 5. Process the smart markers to populate the worksheet.
            designer.Process();

            // 6. Save the result.
            string outputPath = "MergedJsonSmartMarkers.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
