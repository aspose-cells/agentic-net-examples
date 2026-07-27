using System;
using Aspose.Cells;

class ExportActiveWorksheetToJson
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Determine the index of the active worksheet
        int activeSheetIndex = workbook.Worksheets.ActiveSheetIndex;

        // Configure JSON save options to export only the active worksheet
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export only the worksheet whose index matches the active sheet
            SheetIndexes = new int[] { activeSheetIndex }
        };

        // Save the workbook as a JSON file using the default format settings
        workbook.Save("output.json", jsonOptions);
    }
}