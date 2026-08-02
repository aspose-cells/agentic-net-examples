using System;
using Aspose.Cells;
using Aspose.Cells.Json;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some cells with values
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);

            // Add a formula that references the above cells
            worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Configure JSON save options.
            // ExportFormulaValue is not available in this version; the default behavior exports formulas.
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as a JSON file using the configured options
            workbook.Save("WorkbookWithFormulas.json", jsonOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}