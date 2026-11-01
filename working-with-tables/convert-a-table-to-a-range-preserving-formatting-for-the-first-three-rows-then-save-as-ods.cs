// Title: Convert an Excel ListObject to a regular range, preserve the first three rows' formatting, and save as ODS using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert Excel table (ListObject) to range while keeping row styles | How to retain formatting of the first three rows after converting a table to a range with Aspose.Cells | Save a workbook as ODS after converting a ListObject to a normal range in .NET | Reapply cell style after table conversion using Aspose.Cells for C# | Export Excel table to OpenDocument Spreadsheet format with Aspose.Cells
// Tags: listobject range conversion Aspose.Cells | retain first rows formatting Aspose.Cells | export workbook to ODS Aspose.Cells | reapply cell style after table conversion | excel table to range .NET

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an Excel file, checks for a ListObject on the first worksheet, records its area, converts the table to a normal range, explicitly reapplies the existing style to the first three rows to preserve formatting, and then saves the workbook as an ODS file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.ods";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count > 0)
            {
                // Reference the first table in the worksheet
                ListObject table = sheet.ListObjects[0];

                // Store the table's area before conversion
                CellArea tableArea = new CellArea
                {
                    StartRow = table.StartRow,
                    StartColumn = table.StartColumn,
                    EndRow = table.EndRow,
                    EndColumn = table.EndColumn
                };

                // Convert the table to a normal range; formatting of the table cells is retained
                table.ConvertToRange();

                // Preserve formatting explicitly for the first three rows of the former table area
                int rowsToPreserve = Math.Min(3, tableArea.EndRow - tableArea.StartRow + 1);
                for (int r = 0; r < rowsToPreserve; r++)
                {
                    int rowIndex = tableArea.StartRow + r;
                    for (int c = tableArea.StartColumn; c <= tableArea.EndColumn; c++)
                    {
                        // Reapply the existing style to ensure it is kept
                        Style style = sheet.Cells[rowIndex, c].GetStyle();
                        sheet.Cells[rowIndex, c].SetStyle(style);
                    }
                }
            }

            // Save the modified workbook as ODS
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
