// Title: C# – Export Excel to CSV with numbers rounded to 2 decimals using Aspose.Cells
// Description: Loads an XLSX workbook, rounds every numeric cell to two decimal places (MidpointRounding.AwayFromZero), and saves the result as a CSV file—ideal for precise financial reporting.
// Keywords: Aspose.Cells CSV export C# | round numbers to two decimals | financial reporting Excel to CSV | C# round numeric cells Aspose | Save workbook as CSV | US accounting data export | midpoint rounding away from zero
// Common Searches: C# Aspose.Cells round values before CSV export | export Excel to CSV with two‑decimal precision | how to round all numbers in a workbook using Aspose | financial CSV export Aspose.Cells .NET | midpoint rounding for accounting data
// Developer Intent: Apply two‑decimal rounding to every numeric cell in a workbook and generate a CSV file.
// Use Cases: Produce CSV statements for US GAAP‑compliant financial reports. | Create data feeds for accounting software that require fixed‑point monetary values. | Automate monthly ledger exports with uniform two‑decimal rounding.
// AI Prompts: Generate C# code that uses Aspose.Cells to round all numeric cells to two decimal places and save the workbook as CSV. | Explain why MidpointRounding.AwayFromZero is preferred for financial rounding in CSV exports. | Suggest a method to format numbers to two decimals during CSV export without altering the original worksheet values.

using System;
using Aspose.Cells;

// Loads an XLSX workbook, rounds every numeric cell to two decimal places (MidpointRounding.AwayFromZero), and saves the result as a CSV file—ideal for precise financial reporting.
class ExportWorkbookToCsvRounded
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Round all numeric cells to two decimal places
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    double roundedValue = Math.Round(cell.DoubleValue, 2, MidpointRounding.AwayFromZero);
                    cell.PutValue(roundedValue);
                }
            }
        }

        // Export the workbook to CSV format
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
