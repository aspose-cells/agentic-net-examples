// Title: Dynamically Set TableToRangeOptions.LastRow from ListObject Row Count in Aspose.Cells for .NET
// Description: Shows how to read a ListObject's zero‑based last row via the EndRow property, assign that value to TableToRangeOptions.LastRow, and then convert the table into a regular range. The sample creates a workbook, populates data, adds a table, calculates the final row at runtime, and saves the file.
// Keywords: Aspose.Cells | TableToRangeOptions | LastRow | dynamic row count | ListObject | C# | .NET | convert table to range | EndRow property | US developers | global audience
// Common Searches: Aspose.Cells set TableToRangeOptions.LastRow dynamically | Get ListObject last row index in C# | Convert Aspose.Cells table to range using EndRow | How to handle variable row count with TableToRangeOptions | C# example for table-to-range conversion in Aspose.Cells
// Developer Intent: Obtain the table’s final row at runtime and feed it to TableToRangeOptions before converting the ListObject to a range.
// Use Cases: Processing tables whose size changes during execution, ensuring only populated rows are transformed. | Avoiding inclusion of empty rows after adding or removing data from a ListObject. | Exporting a table of unknown dimensions to a range for further manipulation or third‑party integration. | Automating report generation where the number of rows cannot be predetermined.
// AI Prompts: Write C# code that reads a ListObject’s EndRow value and uses it to set TableToRangeOptions.LastRow for ConvertToRange in Aspose.Cells. | Provide an example of converting a variable‑size Aspose.Cells table to a range with dynamic LastRow handling. | Explain step‑by‑step how to retrieve the zero‑based last row of a ListObject and apply it to TableToRangeOptions in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDynamicLastRowDemo
{
    // Shows how to read a ListObject's zero‑based last row via the EndRow property, assign that value to TableToRangeOptions.LastRow, and then convert the table into a regular range. The sample creates a workbook, populates data, adds a table, calculates the final row at runtime, and saves the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (10 rows, 5 columns)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"Data {row}-{col}");
                }
            }

            // Add a ListObject (table) covering the populated range
            int tableIndex = worksheet.ListObjects.Add(0, 0, 9, 4, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Determine the last row index of the table dynamically
            // EndRow returns the zero‑based index of the last row in the table
            int dynamicLastRow = table.EndRow;

            // Create TableToRangeOptions and assign the dynamic LastRow value
            TableToRangeOptions options = new TableToRangeOptions
            {
                LastRow = dynamicLastRow
            };

            // Convert the table to a range using the options
            table.ConvertToRange(options);

            // Save the workbook
            workbook.Save("TableToRange_DynamicLastRow.xlsx");
        }
    }
}
