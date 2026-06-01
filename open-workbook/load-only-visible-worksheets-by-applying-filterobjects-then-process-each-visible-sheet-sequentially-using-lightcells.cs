using System;
using Aspose.Cells;

namespace AsposeCellsVisibleSheetsLightCells
{
    // Custom LoadFilter that loads data only for visible worksheets
    class VisibleSheetLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load full data for visible sheets, otherwise load only structure
            LoadDataFilterOptions = sheet.IsVisible ? LoadDataFilterOptions.All : LoadDataFilterOptions.Structure;
        }
    }

    // LightCellsDataHandler that processes only visible worksheets
    class VisibleSheetHandler : LightCellsDataHandler
    {
        // Called before reading cells of a worksheet
        public bool StartSheet(Worksheet sheet)
        {
            // Process the sheet only if it is visible
            return sheet.IsVisible;
        }

        // Called before processing a row; always process rows
        public bool StartRow(int rowIndex) => true;

        // Called after a row is read; not used here
        public bool ProcessRow(Row row) => true;

        // Called before processing a cell; always process cells
        public bool StartCell(int columnIndex) => true;

        // Called for each cell; here we simply output its address and value
        public bool ProcessCell(Cell cell)
        {
            Console.WriteLine($"Sheet: {cell.Worksheet.Name}, Cell: {cell.Name}, Value: {cell.StringValue}");
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string inputPath = "InputWorkbook.xlsx";

            // Configure load options with the custom filter and LightCells handler
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new VisibleSheetLoadFilter(),
                LightCellsDataHandler = new VisibleSheetHandler()
            };

            // Load the workbook using LightCells mode; only visible sheets' data will be loaded and processed
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // After processing, you may save the workbook if needed
            workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}