using System;
using Aspose.Cells;
using System.IO;

public class LargeExcelWriter
{
    // Creates a large Excel file using LightCells API for efficient memory usage
    public void CreateLargeFile(string outputPath)
    {
        // Initialize an empty workbook (uses the default constructor rule)
        Workbook workbook = new Workbook();

        // Configure OoxmlSaveOptions to use LightCellsDataProvider (uses the Save(string, SaveOptions) rule)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = new CustomLightCellsDataProvider()
        };

        // Save the workbook with the LightCells data provider
        workbook.Save(outputPath, saveOptions);
    }

    // Custom implementation of LightCellsDataProvider that streams data row by row
    private class CustomLightCellsDataProvider : LightCellsDataProvider
    {
        private const int TotalRows = 100_000;   // Example large row count
        private const int TotalCols = 10;        // Example column count per row

        private int currentRow = -1;
        private int currentCol = -1;
        private bool processCurrentSheet = false;

        // Called at the start of each worksheet; we handle only the first sheet (index 0)
        public bool StartSheet(int sheetIndex)
        {
            processCurrentSheet = sheetIndex == 0;
            return processCurrentSheet;
        }

        // Returns the next row index to be written, or -1 when done
        public int NextRow()
        {
            if (!processCurrentSheet) return -1;

            currentRow++;
            currentCol = -1; // reset column index for the new row
            return currentRow < TotalRows ? currentRow : -1;
        }

        // Allows row-level customization (e.g., height); not needed here
        public void StartRow(Row row)
        {
            // No special row handling required
        }

        // Returns the next column index within the current row, or -1 when the row is finished
        public int NextCell()
        {
            currentCol++;
            return currentCol < TotalCols ? currentCol : -1;
        }

        // Populates the current cell with data
        public void StartCell(Cell cell)
        {
            // Example data: "R{row}C{col}"
            cell.PutValue($"R{currentRow}C{currentCol}");
        }

        // Indicates whether string values should be gathered into a global pool for compression
        public bool IsGatherString()
        {
            return true;
        }
    }
}

// Example program entry point
class Program
{
    static void Main()
    {
        var writer = new LargeExcelWriter();
        writer.CreateLargeFile("LargeFile.xlsx");
        Console.WriteLine("Large Excel file created successfully.");
    }
}