// Title: C# – Convert Excel to CSV with Japanese Era Date Formatting using Aspose.Cells
// Description: Load an .xlsx workbook with Aspose.Cells, set the workbook region to Japan, apply the Japanese calendar format "[$-F800]yyyy年m月d日" to every DateTime cell, and save the result as a CSV file so dates appear in the Japanese era style.
// Keywords: Aspose.Cells | C# CSV conversion | Japanese calendar | Japanese era date format | locale Japan | custom date format [$-F800] | .NET Excel export | workbook region Japan | Excel to CSV Aspose | date formatting Japan
// Common Searches: Aspose.Cells export Excel to CSV with Japanese dates | C# set workbook region to Japan for CSV output | How to apply Japanese era format in CSV using Aspose.Cells | Convert Excel file to CSV preserving Japanese calendar | Custom date format [$-F800] in Aspose.Cells CSV
// Developer Intent: Export an Excel workbook to CSV while rendering all date cells in the Japanese era format.
// Use Cases: Create CSV reports for Japanese users where dates follow the era (年/月/日) convention. | Automate batch conversion of multiple workbooks to CSV with Japan‑specific date formatting. | Integrate locale‑aware data export into a .NET pipeline that must comply with Japanese regulatory standards.
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets the region to Japan, applies the Japanese calendar format to all date cells, and saves the workbook as CSV. | Explain the purpose of the format string "[$-F800]yyyy年m月d日" and how Aspose.Cells uses it during CSV export. | Provide a modification to the sample that processes every worksheet and creates separate CSV files while keeping the Japanese date format.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load an .xlsx workbook with Aspose.Cells, set the workbook region to Japan, apply the Japanese calendar format "[$-F800]yyyy年m月d日" to every DateTime cell, and save the result as a CSV file so dates appear in the Japanese era style.
    public class WorkbookToCsvJapaneseCalendar
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Input file not found: {sourcePath}");
                return;
            }

            // Load the workbook (default LoadOptions are sufficient)
            Workbook workbook = new Workbook(sourcePath);

            // Set the workbook's regional settings to Japan.
            // This enables Japanese calendar formatting when applying custom date formats.
            workbook.Settings.Region = CountryCode.Japan;

            // Define a custom date format that uses the Japanese calendar.
            // The format string follows Excel's locale syntax.
            const string japaneseDateFormat = "[$-F800]yyyy年m月d日";

            // Apply the custom format to all cells that contain DateTime values.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through used cells only for efficiency.
            foreach (Cell cell in cells)
            {
                if (cell.Type == CellValueType.IsDateTime)
                {
                    // Get the existing style, modify the custom format, and reapply.
                    Style style = cell.GetStyle();
                    style.Custom = japaneseDateFormat;
                    cell.SetStyle(style);
                }
            }

            // Save the workbook as CSV. The date values will be rendered using the
            // Japanese calendar format defined above.
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook converted to CSV with Japanese date format: {csvPath}");
        }
    }
}
