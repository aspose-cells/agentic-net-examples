using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate the header row with three initial columns
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["C1"].PutValue("Header3");

            // Fill some sample data under the headers
            for (int r = 1; r <= 5; r++)
            {
                cells[r, 0].PutValue($"R{r}C1");
                cells[r, 1].PutValue($"R{r}C2");
                cells[r, 2].PutValue($"R{r}C3");
            }

            // Create a dynamic named range that expands with the number of filled columns in row 1
            // OFFSET starts at A1, height = 1 (header row), width = COUNTA of the entire first row
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name dynamicRange = workbook.Worksheets.Names[nameIndex];
            dynamicRange.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,1,COUNTA(Sheet1!$1:$1))";

            // Verify the initial range size
            AsposeRange range = dynamicRange.GetRange();
            Console.WriteLine($"Initial range address: {range.Address}, columns: {range.ColumnCount}");

            // Insert a new column to the right of the existing data (column D, zero‑based index 3)
            sheet.Cells.InsertColumn(3);
            // Add header and data for the new column
            cells["D1"].PutValue("Header4");
            for (int r = 1; r <= 5; r++)
            {
                cells[r, 3].PutValue($"R{r}C4");
            }

            // Recalculate formulas so the named range reflects the new column count
            workbook.CalculateFormula();

            // Retrieve the updated range and display its new size
            AsposeRange updatedRange = dynamicRange.GetRange();
            Console.WriteLine($"Updated range address: {updatedRange.Address}, columns: {updatedRange.ColumnCount}");

            // Save the workbook (ensure the directory exists)
            string outputPath = "DynamicNamedRange.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}