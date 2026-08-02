// Title: C# – Export All Excel Formulas with Cell Addresses to CSV Using Aspose.Cells
// Description: Load an .xlsx workbook with Aspose.Cells for .NET, iterate every worksheet and populated cell, capture each non‑empty Formula, and write a CSV file that lists the worksheet name, A1‑style cell address, and the escaped formula text.
// Keywords: Aspose.Cells export formulas | C# extract Excel formulas | list cell formulas .NET | save formulas to CSV | worksheet cell address extraction | Aspose.Cells CSV output | Excel formula audit C# | A1 notation formula export | bulk formula extraction Aspose | automate Excel documentation
// Common Searches: how to export Excel formulas to CSV with Aspose.Cells | C# code to list all formulas in a workbook | extract cell formulas and addresses using Aspose.Cells | save Excel formula list as CSV file .NET | Aspose.Cells iterate cells and get formula property
// Developer Intent: Retrieve every formula in a workbook and generate a CSV report containing worksheet name, cell address, and formula text.
// Use Cases: Create an audit trail of all calculated cells for compliance checks. | Migrate spreadsheet logic to another platform by exporting formulas with their locations. | Produce developer documentation that enumerates formulas per worksheet.
// AI Prompts: Generate C# code with Aspose.Cells that exports all formulas to a UTF‑8 CSV, handling quotes correctly. | Adapt the script to export formulas only from a user‑specified worksheet. | Add robust error handling for missing files, permission issues, and log progress to the console.


namespace AsposeCellsFormulaExport
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            string workbookPath = "input.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // Prepare the CSV output file
            string csvPath = "formulas.csv";
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("Worksheet,CellAddress,Formula");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all cells that contain data
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            // Check if the cell has a formula
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                // Get the cell address in A1 style
                                string address = cell.Name;
                                // Escape double quotes in the formula for CSV compliance
                                string formula = cell.Formula.Replace("\"", "\"\"");
                                // Write a CSV line: Worksheet name, cell address, formula
                                writer.WriteLine($"\"{sheet.Name}\",\"{address}\",\"{formula}\"");
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"Formulas have been exported to '{csvPath}'.");
        }
    }
}
