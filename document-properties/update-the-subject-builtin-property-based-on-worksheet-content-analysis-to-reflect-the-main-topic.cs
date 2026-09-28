// Title: Set the Excel workbook Subject built‑in property from the first non‑empty cell in column A using Aspose.Cells for C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, reads the first non‑empty string in column A of the first worksheet, assigns that string to the workbook's Subject built‑in document property, and saves the modified file. | Adapt an existing Aspose.Cells program to automatically determine a worksheet's main topic from column A and update the Subject property of the workbook before saving.
// Common Searches: Aspose.Cells C# set Excel Subject property based on cell value | How to programmatically update built‑in document properties in an Excel file using Aspose.Cells | Read first non‑blank cell in column A and use it as workbook metadata with Aspose.Cells | C# extract topic from first column and assign to Excel Subject built‑in property
// Tags: Aspose.Cells update workbook Subject metadata | extract first column A value C# | modify Excel built‑in document properties programmatically | read first non‑empty string cell Aspose.Cells | automate Excel metadata based on worksheet content

using Aspose.Cells;
using System;
using System.IO;

// // Loads input.xlsx, finds the first non‑empty string in column A of the first worksheet, assigns that string to the workbook's Subject built‑in document property, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Analyze the first worksheet to determine the main topic
            Worksheet sheet = workbook.Worksheets[0];
            string mainTopic = "Untitled";

            // Simple heuristic: first non‑empty string cell in column A
            for (int row = 0; row <= sheet.Cells.MaxDataRow; row++)
            {
                Cell cell = sheet.Cells[row, 0];
                if (cell != null && cell.Type == CellValueType.IsString && !string.IsNullOrWhiteSpace(cell.StringValue))
                {
                    mainTopic = cell.StringValue.Trim();
                    break;
                }
            }

            // Update the Subject built‑in property with the derived topic
            workbook.BuiltInDocumentProperties["Subject"].Value = mainTopic;

            // Output workbook path
            string outputPath = "output.xlsx";

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
