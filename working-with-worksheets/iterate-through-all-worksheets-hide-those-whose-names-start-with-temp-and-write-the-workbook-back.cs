using System;
using Aspose.Cells;

class HideTempSheets
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            // Hide the worksheet if its name starts with "Temp"
            if (sheet.Name.StartsWith("Temp", StringComparison.OrdinalIgnoreCase))
            {
                sheet.IsVisible = false; // or sheet.SetVisible(false, true);
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}