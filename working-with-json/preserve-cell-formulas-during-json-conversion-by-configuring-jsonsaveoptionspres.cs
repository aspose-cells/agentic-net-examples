using System;
using Aspose.Cells;

class PreserveFormulasJsonDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some cells with values
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);

        // Add a formula that references the above cells
        worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Configure JSON save options to preserve formulas in the output
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true
        };

        // Save the workbook as a JSON file using the configured options
        workbook.Save("PreserveFormulasOutput.json", jsonOptions);
    }
}