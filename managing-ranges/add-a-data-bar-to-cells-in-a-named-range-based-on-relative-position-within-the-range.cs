// Title: C# – Add a Data Bar Conditional Formatting to a Named Range using Aspose.Cells
// Description: This example creates a workbook, fills cells A1:A10 with numeric values, defines a named range "MyRange", extracts its coordinates, applies a green data‑bar conditional format with automatic minimum and maximum values to the entire range, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# data bar | conditional formatting named range .NET | add data bar programmatically | named range coordinates Aspose | Excel data bar example C# | Aspose.Cells conditional formatting API | GitHub Aspose.Cells sample
// Common Searches: how to apply data bar to a named range in Aspose.Cells | Aspose.Cells C# data bar conditional formatting example | create and use named ranges with conditional formatting .NET | Aspose.Cells add data bar to cells defined by a name | C# code for data bar conditional format in Excel
// Developer Intent: Apply a data‑bar conditional format to every cell inside a predefined named range.
// Use Cases: Show progress bars for task completion percentages stored in a named range. | Highlight monthly sales figures with green data bars to compare performance. | Build a KPI dashboard where data‑bars automatically adjust as the underlying named range updates.
// AI Prompts: Generate C# code with Aspose.Cells to add a red data bar (min 0, max 100) to the named range B2:B20. | Explain how to retrieve a named range’s start row, end row, start column, and end column for use in a CellArea. | Provide a sample that applies three conditional formats—including a data bar—to the same named range in a workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDataBarNamedRange
{
    // This example creates a workbook, fills cells A1:A10 with numeric values, defines a named range "MyRange", extracts its coordinates, applies a green data‑bar conditional format with automatic minimum and maximum values to the entire range, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in column A (rows 1-10)
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 0].PutValue(i + 1);
                }

                // Define a named range that covers the populated cells (A1:A10)
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$A$10";

                // Retrieve the Range object for the named range
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                AsposeRange range = namedRange.GetRange();

                // Add a new conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

                // Define the area for the conditional formatting using the range's coordinates
                CellArea area = new CellArea
                {
                    StartRow = range.FirstRow,
                    EndRow = range.FirstRow + range.RowCount - 1,
                    StartColumn = range.FirstColumn,
                    EndColumn = range.FirstColumn + range.ColumnCount - 1
                };
                cfCollection.AddArea(area);

                // Add a DataBar condition
                int conditionIdx = cfCollection.AddCondition(FormatConditionType.DataBar);
                FormatCondition condition = cfCollection[conditionIdx];

                // Configure the DataBar (automatic min/max, green color)
                DataBar dataBar = condition.DataBar;
                dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
                dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
                dataBar.Color = Color.Green;
                dataBar.ShowValue = true;

                // Save the workbook
                string outputPath = "DataBar_NamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
