using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one sparkline group already present
        if (worksheet.SparklineGroups.Count > 0)
        {
            // Retrieve the first existing sparkline group
            SparklineGroup sparklineGroup = worksheet.SparklineGroups[0];

            // Define the data range for the new sparkline (e.g., cells A1:D1 of the same sheet)
            string dataRange = worksheet.Name + "!A1:D1";

            // Specify the location where the new sparkline will be placed
            // Row and column are zero‑based indexes; here we place it in cell F1 (row 0, column 5)
            int targetRow = 0;
            int targetColumn = 5;

            // Add the new sparkline to the existing group
            sparklineGroup.Sparklines.Add(dataRange, targetRow, targetColumn);
        }

        // Save the workbook with the newly added sparkline
        workbook.Save("output.xlsx");
    }
}