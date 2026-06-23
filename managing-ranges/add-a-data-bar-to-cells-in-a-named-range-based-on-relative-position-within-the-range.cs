using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace AsposeCellsDataBarNamedRange
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in columns A to C (rows 1-10)
                for (int row = 0; row < 10; row++)
                {
                    sheet.Cells[row, 0].PutValue(row + 1);          // Column A
                    sheet.Cells[row, 1].PutValue((row + 1) * 10);   // Column B
                    sheet.Cells[row, 2].PutValue((row + 1) * 5);    // Column C
                }

                // Define a named range that covers the populated cells (A1:C10)
                int nameIndex = sheet.Workbook.Worksheets.Names.Add("MyRange");
                sheet.Workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$C$10";

                // Retrieve the Range object for the named range
                Name namedRange = sheet.Workbook.Worksheets.Names["MyRange"];
                AsposeRange range = namedRange.GetRange();

                // Build a CellArea that matches the range (required for conditional formatting)
                CellArea area = new CellArea
                {
                    StartRow = range.FirstRow,
                    StartColumn = range.FirstColumn,
                    EndRow = range.FirstRow + range.RowCount - 1,
                    EndColumn = range.FirstColumn + range.ColumnCount - 1
                };

                // Add a new conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

                // Apply the area of the named range to this conditional formatting
                fcs.AddArea(area);

                // Add a DataBar condition
                int conditionIdx = fcs.AddCondition(FormatConditionType.DataBar);
                FormatCondition condition = fcs[conditionIdx];

                // Configure the DataBar (automatic min/max, green color, hide numeric values)
                DataBar dataBar = condition.DataBar;
                dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
                dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
                dataBar.Color = Color.Green;
                dataBar.ShowValue = false;

                // Save the workbook
                string outputPath = "DataBar_NamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}