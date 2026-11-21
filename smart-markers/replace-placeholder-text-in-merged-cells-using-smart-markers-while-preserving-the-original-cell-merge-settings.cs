// Title: Replace a smart‑marker placeholder in merged Excel cells while preserving the merge using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells replace smart marker in a merged range without losing merge | how to keep merged cells when processing smart markers with WorkbookDesigner | replace ${Name} placeholder in merged cells A1:C1 using Aspose.Cells | Aspose.Cells smart markers preserve cell merge after processing | example of merging cells and using smart markers in a C# workbook
// Tags: smart marker replacement merged cells Aspose.Cells | WorkbookDesigner preserve merge settings | C# merge Excel range with smart marker | Aspose.Cells process smart markers in merged range | replace placeholder in merged Excel cells C#

using System;
using System.Data;
using Aspose.Cells;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The sample creates a workbook, merges cells A1:C1, inserts the smart‑marker ${Name} into the merged range, binds a DataTable with a Name column as the data source, processes the smart markers with WorkbookDesigner while retaining the original merge, and saves the result to Output.xlsx.
class ReplaceSmartMarkerInMergedCells
{
    static void Main()
    {
        try
        {
            // Create a new workbook (no template file required)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create and merge the range A1:C1
            AsposeRange mergedRange = sheet.Cells.CreateRange("A1", "C1");
            mergedRange.Merge();

            // Insert a smart marker placeholder into the merged cells
            sheet.Cells["A1"].PutValue("${Name}");

            // Prepare a simple data source
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Rows.Add("John Doe");

            // Use WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process(); // merged cells retain their merge settings

            // Define output path and ensure the directory exists
            string outputPath = "Output.xlsx";
            try
            {
                string directory = System.IO.Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !System.IO.Directory.Exists(directory))
                {
                    System.IO.Directory.CreateDirectory(directory);
                }

                // Save the result
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
