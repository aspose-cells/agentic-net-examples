using System;
using Aspose.Cells;

class CsvPrecisionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Configure load options to keep full numeric precision
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
        loadOptions.Separator = ',';               // CSV delimiter
        loadOptions.ConvertNumericData = true;     // Convert numeric strings to numbers
        loadOptions.KeepPrecision = true;         // Preserve precision, avoid exponential notation

        // Path to the CSV file (replace with your actual file path)
        string csvPath = "input.csv";

        // Import the CSV data starting at cell A1 (row 0, column 0)
        workbook.Worksheets[0].Cells.ImportCSV(csvPath, loadOptions, 0, 0);

        // Ensure output uses the highest precision format (G17)
        CellsHelper.SignificantDigitsType = SignificantDigitsType.G17;

        // Apply a custom number format to all used cells to prevent scientific notation
        Style style = workbook.CreateStyle();
        style.Custom = "0.##############################"; // up to 30 decimal places
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the style to the used range of the worksheet
        Aspose.Cells.Range usedRange = workbook.Worksheets[0].Cells.MaxDisplayRange;
        usedRange.ApplyStyle(style, flag);

        // Save the workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}