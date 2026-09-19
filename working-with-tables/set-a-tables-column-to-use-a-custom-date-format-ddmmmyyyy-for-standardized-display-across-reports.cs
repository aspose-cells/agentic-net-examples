// Title: Apply a custom date format dd-MMM-yyyy to a specific table column using Aspose.Cells for .NET (C#)
// AI Prompts: Create a Style with Custom = "dd-MMM-yyyy" and assign it to the data range of a ListObject column in a C# Aspose.Cells workbook. | Load an existing Excel file (or instantiate a new Workbook), ensure a table exists, then set the second column of the first table to display dates using the dd-MMM-yyyy pattern. | Save the workbook after applying the custom date style to the selected table column and confirm the output file is generated.
// Common Searches: Aspose.Cells C# format ListObject column with custom date pattern | set dd-MMM-yyyy date style for Excel table column using Aspose.Cells .NET | how to apply custom date format to a table column in Aspose.Cells workbook | C# Aspose.Cells change date display format for specific column in a table | apply custom number format to table column Aspose.Cells example
// Tags: format table column date style Aspose.Cells C# | custom date pattern dd-MMM-yyyy Aspose.Cells | listobject column styling .NET | apply custom number format to Excel table column | c# workbook column date formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The program loads or creates an Excel workbook, ensures a table exists, and applies a custom date format (dd-MMM-yyyy) to the second column of the first table using Aspose.Cells for .NET, then saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    // Add a default worksheet to avoid index errors.
                    workbook.Worksheets.Add("Sheet1");
                }

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table.
                if (worksheet.ListObjects.Count == 0)
                {
                    // Create a sample table for demonstration if none exists.
                    // This creates a table covering A1:C5 with a header row.
                    AsposeRange dataRange = worksheet.Cells.CreateRange("A1:C5");
                    dataRange[0, 0].PutValue("Header1");
                    dataRange[0, 1].PutValue("Header2");
                    dataRange[0, 2].PutValue("Header3");

                    // Add the table using the overload that specifies the range dimensions.
                    worksheet.ListObjects.Add(
                        dataRange.FirstRow,
                        dataRange.FirstColumn,
                        dataRange.RowCount,
                        dataRange.ColumnCount,
                        true);
                }

                // Retrieve the first table (ListObject) on the worksheet.
                ListObject table = worksheet.ListObjects[0];

                // Column index to format (0‑based). Example: second column in the table.
                int columnIndex = 1;

                // Apply a default style to the header cell of the selected column.
                int headerRow = table.StartRow;
                int headerCol = table.StartColumn + columnIndex;
                worksheet.Cells[headerRow, headerCol].SetStyle(workbook.CreateStyle());

                // Determine the data range for the selected column (excluding the header).
                int dataStartRow = table.StartRow + 1; // first data row
                int dataEndRow = table.EndRow;         // last data row
                int dataColumn = table.StartColumn + columnIndex;
                int rowCount = dataEndRow - dataStartRow + 1;

                // Create a range that represents the column's data cells.
                AsposeRange columnDataRange = worksheet.Cells.CreateRange(dataStartRow, dataColumn, rowCount, 1);

                // Create a style with the desired date format.
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14;               // Built‑in date style (optional)
                dateStyle.Custom = "dd-MMM-yyyy";

                // Apply the style to the whole column range (data part).
                columnDataRange.SetStyle(dateStyle);

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the updated formatting.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
