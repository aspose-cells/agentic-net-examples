// Title: Set custom date format 'dd‑MMM‑yyyy' for a table column using Aspose.Cells .NET
// Description: Demonstrates how to create a workbook, add a ListObject (Excel table), define a style with the custom date pattern dd‑MMM‑yyyy, apply the style to the entire column of the table, and save the file.
// Keywords: Aspose.Cells | C# | custom date format | dd-MMM-yyyy | ListObject | table column formatting | Excel date style | SetStyle | Range styling | Excel reporting
// Common Searches: Aspose.Cells format date column C# | apply custom date style to ListObject column | dd-MMM-yyyy format in Aspose.Cells workbook | how to set table column style Aspose.Cells .NET | standardize date display in Excel with Aspose.Cells
// Developer Intent: Apply a uniform dd‑MMM‑yyyy date format to a specific table column in an Excel file generated with Aspose.Cells.
// Use Cases: Generating financial reports where all dates must appear as 15‑Jan‑2023. | Ensuring consistent date presentation across multiple tables in automated Excel exports. | Reusing a predefined date style for dashboards that consume Excel data.
// AI Prompts: Show C# code that creates a ListObject and sets its first column to the 'dd-MMM-yyyy' date format with Aspose.Cells. | Explain how to define a custom date style once and apply it to several table columns in an existing workbook. | Provide steps to reuse a date style for multiple tables when building Excel reports using Aspose.Cells .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // Demonstrates how to create a workbook, add a ListObject (Excel table), define a style with the custom date pattern dd‑MMM‑yyyy, apply the style to the entire column of the table, and save the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data: a header and two date values
                worksheet.Cells["A1"].PutValue("Date");
                worksheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
                worksheet.Cells["A3"].PutValue(new DateTime(2023, 2, 20));

                // Add a table (ListObject) that includes the header and the two data rows
                // Parameters: firstRow, firstColumn, totalRows (excluding header), totalColumns, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Create a style with the desired custom date format
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "dd-MMM-yyyy";

                // Apply the style to the entire column of the table (header + data rows)
                table.ListColumns[0].Range.SetStyle(dateStyle);

                // Determine output path and ensure directory exists
                string outputFile = "TableDateFormat.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
