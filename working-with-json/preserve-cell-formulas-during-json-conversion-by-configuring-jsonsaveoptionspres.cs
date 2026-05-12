using System;
using Aspose.Cells;

<<<<<<< HEAD
class Program
=======
class PreserveFormulasJsonDemo
>>>>>>> 36c4d37ac3162a2d03072975974d49755a657950
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

<<<<<<< HEAD
        // Configure JSON save options (formulas are exported by default)
        JsonSaveOptions jsonOptions = new JsonSaveOptions();

        // Save the workbook as a JSON file using the configured options
        workbook.Save("WorkbookWithFormulas.json", jsonOptions);

        Console.WriteLine("Workbook saved to JSON with formulas preserved.");
=======
        // Configure JSON save options to preserve formulas in the output
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true
        };

        // Save the workbook as a JSON file using the configured options
        workbook.Save("PreserveFormulasOutput.json", jsonOptions);
>>>>>>> 36c4d37ac3162a2d03072975974d49755a657950
    }
}