// Title: Filter out rows with Status = 'Inactive' using Aspose.Cells AutoFilter and export the remaining data to JSON in C#
// Common Searches: C# Aspose.Cells filter rows where Status column is Inactive before JSON export | How to apply a custom AutoFilter to exclude specific values and save as JSON with Aspose.Cells | Export filtered Excel range to JSON using Aspose.Cells .NET API | Locate column by header name and apply NotEqual filter in Aspose.Cells | Create JSON from filtered worksheet range in C# Aspose.Cells
// Tags: custom autofilter not equal Aspose.Cells | export filtered range to JSON C# | locate column by header Aspose.Cells | jsonsaveoptions toexcelstruct false Aspose.Cells | worksheet autofilter range creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Json;

namespace AsposeCellsJsonExport
{
    // The example loads an Excel workbook, finds the column titled "Status", defines an AutoFilter covering all used cells, applies a custom NotEqual filter to exclude rows where Status equals "Inactive", refreshes the filter, creates a Range object from the filtered area, converts that range to JSON using JsonSaveOptions with ToExcelStruct set to false, and writes the resulting JSON to an output file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputJsonPath = "output.json";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Locate the column that contains the header "Status"
                // ------------------------------------------------------------
                const string statusHeader = "Status";
                int statusColumnIndex = -1;
                int headerRowIndex = 0; // assuming the first row contains headers

                int maxColumn = worksheet.Cells.MaxDataColumn;
                for (int col = 0; col <= maxColumn; col++)
                {
                    string cellValue = worksheet.Cells[headerRowIndex, col].StringValue?.Trim();
                    if (string.Equals(cellValue, statusHeader, StringComparison.OrdinalIgnoreCase))
                    {
                        statusColumnIndex = col;
                        break;
                    }
                }

                if (statusColumnIndex == -1)
                {
                    Console.WriteLine($"Error: Could not find a column with header \"{statusHeader}\".");
                    return;
                }

                // ------------------------------------------------------------
                // 2. Define the autofilter range (including all used rows/columns)
                // ------------------------------------------------------------
                int firstDataRow = headerRowIndex + 1;
                int lastDataRow = worksheet.Cells.MaxDataRow;
                int lastColumn = worksheet.Cells.MaxDataColumn;

                string startAddress = CellsHelper.CellIndexToName(headerRowIndex, 0);
                string endAddress = CellsHelper.CellIndexToName(lastDataRow, lastColumn);
                string filterRange = $"{startAddress}:{endAddress}";

                worksheet.AutoFilter.Range = filterRange;

                // ------------------------------------------------------------
                // 3. Apply a custom filter to exclude rows where Status = "Inactive"
                // ------------------------------------------------------------
                // Note: the correct overload is Custom(int columnIndex, FilterOperatorType filterOperator, string criteria)
                worksheet.AutoFilter.Custom(statusColumnIndex, FilterOperatorType.NotEqual, "Inactive");
                worksheet.AutoFilter.Refresh();

                // ------------------------------------------------------------
                // 4. Export the filtered range to JSON
                // ------------------------------------------------------------
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    // Export as a simple array of objects (not Excel structure)
                    ToExcelStruct = false
                    // Column names are exported by default; no explicit property needed
                };

                // Create a Range object from the autofilter address
                Aspose.Cells.Range exportRange = worksheet.Cells.CreateRange(worksheet.AutoFilter.Range);

                // Convert the range to JSON
                string jsonResult = exportRange.ToJson(jsonOptions);

                // Write JSON to file
                File.WriteAllText(outputJsonPath, jsonResult);
                Console.WriteLine($"JSON export completed successfully. Output written to \"{outputJsonPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
