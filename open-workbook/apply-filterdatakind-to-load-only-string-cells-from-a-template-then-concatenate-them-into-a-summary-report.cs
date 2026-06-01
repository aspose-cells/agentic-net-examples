using System;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Paths to the template and the output summary workbook
        string templatePath = "template.xlsx";
        string outputPath = "summary.xlsx";

        // Configure load options to load only cells that contain string values
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellString);

        // Load the workbook from the template using the specified filter
        Workbook sourceWorkbook = new Workbook(templatePath, loadOptions);

        // StringBuilder to accumulate all string values from the loaded workbook
        StringBuilder concatenatedStrings = new StringBuilder();

        // Iterate through each worksheet and its used cells
        foreach (Worksheet sheet in sourceWorkbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    // After applying the filter, only string cells are loaded,
                    // but we still verify the cell type for safety.
                    if (cell != null && cell.Type == CellValueType.IsString)
                    {
                        concatenatedStrings.Append(cell.StringValue);
                        concatenatedStrings.Append(' '); // separator between values
                    }
                }
            }
        }

        // Create a new workbook that will hold the summary report
        Workbook summaryWorkbook = new Workbook();
        Worksheet summarySheet = summaryWorkbook.Worksheets[0];
        summarySheet.Name = "Summary";

        // Write the concatenated string into cell A1 of the summary sheet
        summarySheet.Cells["A1"].PutValue(concatenatedStrings.ToString().Trim());

        // Save the summary workbook to the specified output path
        summaryWorkbook.Save(outputPath);
    }
}