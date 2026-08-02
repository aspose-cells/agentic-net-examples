// Title: Convert QueryTable‑Linked ListObject to a Standard Table with Aspose.Cells for .NET
// Description: A C# example that loads a workbook, detects ListObjects attached to QueryTables, detaches the QueryTable while keeping the data, and recreates the range as a regular ListObject preserving the original table style. The modified workbook is then saved as a new file.
// Keywords: Aspose.Cells | C# | QueryTable conversion | ListObject to range | preserve table style | remove external data connection | ConvertToRange | Excel table manipulation
// Common Searches: Aspose.Cells convert QueryTable to regular table | C# detach QueryTable from ListObject | preserve formatting when removing QueryTable | replace QueryTable with standard table Aspose.Cells | how to use ConvertToRange in Aspose.Cells
// Developer Intent: Replace any ListObject linked to a QueryTable with a static worksheet table while retaining the cell data and original table formatting.
// Use Cases: Transform dynamic query‑driven tables into static tables before publishing a workbook. | Eliminate external data connections for security or compatibility while keeping visual layout. | Prepare Excel files for environments that do not support QueryTables, such as older Office versions or third‑party viewers.
// AI Prompts: Write C# code using Aspose.Cells that scans a worksheet for ListObjects with QueryTables, calls ConvertToRange, and recreates them as regular ListObjects preserving the TableStyleName. | Show how to detach a QueryTable from a ListObject, keep its data, and add a new static table with the same range and style in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsQueryTableConversion
{
    // A C# example that loads a workbook, detects ListObjects attached to QueryTables, detaches the QueryTable while keeping the data, and recreates the range as a regular ListObject preserving the original table style. The modified workbook is then saved as a new file.
    class Program
    {
        static void Main()
        {
            const string inputFile = "QueryTableWorkbook.xlsx";
            const string outputFile = "ConvertedWorkbook.xlsx";

            try
            {
                // Ensure the input workbook exists; create a simple one if it does not.
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found. Creating a sample workbook.");

                    Workbook sampleWb = new Workbook();
                    Worksheet sampleWs = sampleWb.Worksheets[0];

                    // Add sample data.
                    sampleWs.Cells["A1"].PutValue("Header1");
                    sampleWs.Cells["B1"].PutValue("Header2");
                    sampleWs.Cells["A2"].PutValue("Data1");
                    sampleWs.Cells["B2"].PutValue("Data2");

                    // Add a regular ListObject (no QueryTable) for demonstration.
                    int loIdx = sampleWs.ListObjects.Add(0, 0, 1, 1, true);
                    sampleWs.ListObjects[loIdx].TableStyleName = "TableStyleMedium2";

                    sampleWb.Save(inputFile);
                }

                // Load the workbook that may contain a QueryTable.
                Workbook workbook = new Workbook(inputFile);
                Worksheet sheet = workbook.Worksheets[0];

                // Iterate through ListObjects in reverse order.
                for (int i = sheet.ListObjects.Count - 1; i >= 0; i--)
                {
                    ListObject listObj = sheet.ListObjects[i];

                    // If the ListObject is linked to a QueryTable, convert it.
                    if (listObj.QueryTable != null)
                    {
                        int startRow = listObj.StartRow;
                        int startColumn = listObj.StartColumn;
                        int endRow = listObj.EndRow;
                        int endColumn = listObj.EndColumn;

                        // Remove the QueryTable association while preserving data.
                        listObj.ConvertToRange();

                        // Add a new regular ListObject on the same range.
                        int newIndex = sheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
                        ListObject newTable = sheet.ListObjects[newIndex];

                        // Preserve the original table style.
                        newTable.TableStyleName = listObj.TableStyleName;
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
