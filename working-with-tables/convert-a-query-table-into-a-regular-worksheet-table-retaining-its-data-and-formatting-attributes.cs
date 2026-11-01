// Title: Convert an Excel QueryTable to a regular ListObject with data and header formatting preserved using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook, finds the first QueryTable, creates a ListObject over its result range, copies the header cell style to the new table, removes the original QueryTable, and saves the workbook. | Write a method with Aspose.Cells that transforms a QueryTable into a standard worksheet table while keeping the original data and header formatting intact.
// Common Searches: asp.net aspose.cells replace querytable with listobject preserving formatting | c# convert Excel query table to regular table using Aspose.Cells | how to copy header style when converting querytable to listobject in Aspose.Cells | remove querytable after creating listobject Aspose.Cells C# example
// Tags: Aspose.Cells QueryTable to ListObject conversion | header formatting retention Aspose.Cells | C# replace Excel query table with worksheet table | delete original QueryTable after conversion | copy result range into ListObject

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program loads an Excel workbook, locates the first QueryTable, creates a ListObject (regular worksheet table) over the same cell range, copies the header cell style to the new table, optionally deletes the original QueryTable, and saves the updated workbook to the specified output path.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputFilePath = "{inputFilePath}";
        string outputFilePath = "{outputFilePath}";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"Input file not found: {inputFilePath}");

            // Load the workbook that contains the query table
            Workbook workbook = new Workbook(inputFilePath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one query table on the sheet
            if (sheet.QueryTables.Count > 0)
            {
                // Get the first query table (or locate by name if required)
                QueryTable queryTable = sheet.QueryTables[0];

                // Obtain the range that holds the query table's result data
                Aspose.Cells.Range resultRange = queryTable.ResultRange;

                // Determine the boundaries of the range
                int firstRow = resultRange.FirstRow;
                int firstColumn = resultRange.FirstColumn;
                int totalRows = resultRange.RowCount;
                int totalColumns = resultRange.ColumnCount;

                // Add a regular worksheet table (ListObject) over the same cells
                // ListObjects.Add returns the index of the newly added ListObject
                int listObjectIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
                ListObject listObject = sheet.ListObjects[listObjectIndex];

                // Set a display name for the new table
                if (!string.IsNullOrEmpty(queryTable.Name))
                {
                    listObject.DisplayName = queryTable.Name + "_Table";
                }

                // If the ListObject shows a header row, copy the header cell style
                if (listObject.ShowHeaderRow)
                {
                    Style headerStyle = sheet.Cells[firstRow, firstColumn].GetStyle();
                    for (int col = 0; col < totalColumns; col++)
                    {
                        sheet.Cells[firstRow, firstColumn + col].SetStyle(headerStyle);
                    }
                }

                // Optionally remove the original query table definition
                sheet.QueryTables.RemoveAt(0);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the converted regular table
            workbook.Save(outputFilePath);
            Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
