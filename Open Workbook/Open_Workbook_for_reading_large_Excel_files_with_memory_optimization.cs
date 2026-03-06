using System;
using Aspose.Cells;

class LargeExcelReader
{
    public static void Run()
    {
        // Path to the large Excel file
        string filePath = "LargeFile.xlsx";

        // Create load options with memory‑optimized settings
        LoadOptions loadOptions = new LoadOptions
        {
            // Use compact memory mode to reduce RAM consumption
            MemorySetting = MemorySetting.MemoryPreference,
            // Skip formula parsing if formulas are not needed
            ParsingFormulaOnOpen = false
        };

        // Load the workbook using the memory‑optimized options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Iterate through used rows and columns sequentially (efficient for large files)
        int maxRow = sheet.Cells.MaxDataRow;
        int maxCol = sheet.Cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                var cell = sheet.Cells[row, col];
                // Example processing: output cell address and value
                Console.WriteLine($"R{row + 1}C{col + 1}: {cell.StringValue}");
            }
        }

        // If you need to save after processing, uncomment the line below
        // workbook.Save("Processed.xlsx", SaveFormat.Xlsx);
    }
}

class Program
{
    static void Main(string[] args)
    {
        LargeExcelReader.Run();
    }
}