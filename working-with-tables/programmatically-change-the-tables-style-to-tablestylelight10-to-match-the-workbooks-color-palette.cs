// Title: How to apply the built‑in TableStyleLight10 to an Excel table using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change existing ListObject style to TableStyleLight10 | set table style programmatically in .xlsx with Aspose.Cells .NET | apply built‑in table style to Excel table via C# Aspose.Cells API | update ListObject TableStyleType enum in C# Aspose.Cells example | change Excel table formatting to TableStyleLight10 using Aspose.Cells
// Tags: Aspose.Cells set ListObject style | TableStyleLight10 usage Aspose.Cells | programmatic Excel table styling .NET | C# update table style enum | apply built‑in table style to worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, retrieves a ListObject named "Table1" from the first worksheet, parses the string "TableStyleLight10" into the TableStyleType enum, assigns this style to the table, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string tableName = "Table1";
            const string styleName = "TableStyleLight10";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index or name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the table (ListObject) by name
                ListObject table = worksheet.ListObjects[tableName];
                if (table == null)
                {
                    Console.WriteLine($"Error: Table \"{tableName}\" not found in the worksheet.");
                    return;
                }

                // Parse the style name to the corresponding enum value
                if (!Enum.TryParse<TableStyleType>(styleName, out TableStyleType styleType))
                {
                    Console.WriteLine($"Error: Table style \"{styleName}\" is not a valid built‑in style.");
                    return;
                }

                // Apply the desired table style
                table.TableStyleType = styleType;

                // Save the workbook with the updated table style
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
