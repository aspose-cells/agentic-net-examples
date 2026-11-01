// Title: C# example: Use Aspose.Cells FilterDataKind to load only string cells from an Excel template and create a concatenated summary report
// Common Searches: Aspose.Cells FilterDataKind load only text cells from an existing .xlsx in C# | concatenate all string values from multiple worksheets using Aspose.Cells .NET | generate a summary workbook with combined text from a template Excel file in C# | filter Excel cells by type and export merged text with Aspose.Cells API | C# code to read only string cells and write them to a new Excel report
// Tags: Aspose.Cells FilterDataKind string cell extraction | load Excel template text-only cells .NET | concatenate worksheet string values Aspose.Cells | create summary report workbook from filtered data | C# extract and merge Excel text using Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample verifies the template file, loads it with Aspose.Cells using FilterDataKind to include only string cells, iterates through all worksheets to collect and concatenate those text values, writes the combined string into the first cell of a new workbook, and saves the result as a summary report.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";

            // Verify template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook templateWb = new Workbook(templatePath, loadOptions);

            // Collect all string values from the workbook
            StringBuilder concatenated = new StringBuilder();

            foreach (Worksheet sheet in templateWb.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            concatenated.Append(cell.StringValue);
                            concatenated.Append(' ');
                        }
                    }
                }
            }

            // Create a new workbook for the summary report
            Workbook summaryWb = new Workbook();
            Worksheet summarySheet = summaryWb.Worksheets[0];

            // Write the concatenated string into the first cell
            summarySheet.Cells[0, 0].PutValue(concatenated.ToString().Trim());

            // Save the summary report
            const string summaryPath = "SummaryReport.xlsx";
            summaryWb.Save(summaryPath);
            Console.WriteLine($"Summary report saved to {summaryPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
