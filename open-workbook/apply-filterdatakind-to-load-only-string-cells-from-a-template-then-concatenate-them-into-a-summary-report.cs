// Title: Extract only string cells from an Excel template and concatenate them into a summary workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to read only text cells from an Excel template using Aspose.Cells in C# | concatenate all string values from multiple worksheets into a single cell with Aspose.Cells .NET | create a summary Excel file by merging text cells from a template workbook in C# | Aspose.Cells example for extracting string values and generating a combined report | C# code to iterate used range and collect string cells with Aspose.Cells
// Tags: collect string values Aspose.Cells | concatenate text cells Aspose.Cells | generate summary workbook Aspose.Cells | iterate used range Aspose.Cells | write concatenated text Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The C# program checks for a template Excel file, loads it with Aspose.Cells, iterates through each worksheet's used range, gathers cells whose type is string, concatenates their text separated by spaces, creates a new workbook named "Summary", writes the combined string into the first cell, saves the file as SummaryReport.xlsx, and includes basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";

            // Ensure the template file exists before loading
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file '{templatePath}' not found.");
                return;
            }

            // Load the workbook (no filtering needed for string cells)
            Workbook templateWb = new Workbook(templatePath);

            // Collect all string values from the workbook
            StringBuilder sb = new StringBuilder();
            foreach (Worksheet sheet in templateWb.Worksheets)
            {
                Cells cells = sheet.Cells;
                Aspose.Cells.Range usedRange = cells.MaxDisplayRange;

                int startRow = usedRange.FirstRow;
                int endRow = startRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = startCol + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell != null && cell.Type == CellValueType.IsString)
                        {
                            sb.Append(cell.StringValue);
                            sb.Append(' ');
                        }
                    }
                }
            }

            // Create a new workbook for the summary report
            Workbook summaryWb = new Workbook();
            Worksheet summarySheet = summaryWb.Worksheets[0];
            summarySheet.Name = "Summary";

            // Write the concatenated string into the first cell
            summarySheet.Cells[0, 0].PutValue(sb.ToString().Trim());

            // Save the summary report
            const string summaryPath = "SummaryReport.xlsx";
            summaryWb.Save(summaryPath, SaveFormat.Xlsx);
            Console.WriteLine($"Summary report saved to '{summaryPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
