// Title: Convert an Excel workbook to CSV while applying Japanese era (wareki) date formatting with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, sets the workbook CultureInfo to Japanese (ja-JP), formats every DateTime cell with the Japanese era pattern, and saves the first worksheet as a CSV file. | Show how to iterate over all cells in an Aspose.Cells workbook, detect cells of type DateTime, assign a wareki date style, and export the sheet to CSV in a .NET application.
// Common Searches: aspocells export to csv with Japanese era date format c# | c# set workbook culture to ja-jp before csv conversion using aspose.cells | apply wareki date formatting to all cells when saving excel as csv | how to convert xlsx to csv preserving Japanese calendar in .net | aspocells custom number format for Japanese calendar during csv export
// Tags: Japanese era custom format Aspose.Cells | export first worksheet to CSV Aspose.Cells C# | set workbook CultureInfo ja-JP Aspose.Cells | detect DateTime cell type Aspose.Cells | apply wareki number format ggge年M月d日

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Loads input.xlsx, sets the workbook culture to Japanese (ja-JP), applies the Japanese era format "ggge年M月d日" to every DateTime cell across all worksheets, and saves the first sheet as output.csv in CSV format using Aspose.Cells.
class WorkbookToCsvJapaneseCalendar
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook culture to Japanese (Japan) for Japanese calendar handling
            workbook.Settings.CultureInfo = new CultureInfo("ja-JP");

            // Custom number format using the Japanese era (wareki), e.g., "令和3年5月21日"
            string japaneseDateFormat = "ggge年M月d日";

            // Iterate through all worksheets and apply the custom Japanese date format to date cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                // Determine start and end indices
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startColumn = usedRange.FirstColumn;
                int endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Apply format only to cells containing DateTime values
                        if (cell.Type == CellValueType.IsDateTime)
                        {
                            Style style = cell.GetStyle();
                            style.Custom = japaneseDateFormat;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Output CSV path
            string outputPath = "output.csv";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the first worksheet as CSV (default behavior)
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook successfully saved as CSV to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
