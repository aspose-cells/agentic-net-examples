// Title: Extract only string cells from an Excel template with Aspose.Cells, concatenate them, and save as a summary report workbook (C#)
// AI Prompts: Write C# code that opens a .xlsx template using Aspose.Cells, iterates through every worksheet, collects cells where Cell.Type == CellValueType.IsString, concatenates the text values with spaces, and writes the combined string to cell A1 of a new workbook. | Create a console application that accepts the template file path and the output report path as command‑line arguments, uses Aspose.Cells to load the template, extracts only text cells, builds a single summary string, and saves it as an .xlsx file. | Adapt the example to filter numeric cells instead of strings, calculate their sum, and place the total in the summary workbook.
// Common Searches: Aspose.Cells C# read only text cells from an existing workbook | how to concatenate all string values from multiple worksheets using Aspose.Cells | generate a summary Excel file from a template by extracting string cells in .NET | filter cells by type when loading an Excel file with Aspose.Cells C# | combine text from every cell in a workbook into one cell using Aspose.Cells
// Tags: extract string cells Aspose.Cells | concatenate worksheet text values C# | generate summary report workbook Aspose.Cells | filter cells by CellValueType Aspose.Cells | save concatenated string to cell A1 Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads Template.xlsx with Aspose.Cells, iterates through all worksheets' used ranges, selects cells whose Type is IsString, concatenates their text values separated by spaces, writes the resulting string to cell A1 of a new workbook, and saves the file as SummaryReport.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the template workbook
            string templatePath = "Template.xlsx";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the workbook (no data filter – we will check cell types manually)
            Workbook templateWorkbook;
            try
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                templateWorkbook = new Workbook(templatePath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load template workbook: {ex.Message}");
                return;
            }

            // StringBuilder to accumulate all string values
            StringBuilder concatenated = new StringBuilder();

            // Iterate through all worksheets
            foreach (Worksheet sheet in templateWorkbook.Worksheets)
            {
                // Get the used range of the sheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                if (usedRange == null) continue;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell != null && cell.Type == CellValueType.IsString)
                        {
                            concatenated.Append(cell.StringValue);
                            concatenated.Append(' ');
                        }
                    }
                }
            }

            // Create a new workbook for the summary report
            Workbook summaryWorkbook = new Workbook();
            Worksheet summarySheet = summaryWorkbook.Worksheets[0];

            // Write the concatenated string into cell A1
            summarySheet.Cells["A1"].PutValue(concatenated.ToString().Trim());

            // Save the summary report
            string outputPath = "SummaryReport.xlsx";
            try
            {
                summaryWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Summary report saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save summary report: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
