using System;
using Aspose.Cells;

class LightCellsPdfConversion
{
    static void Main()
    {
        // Path to the large Excel file
        string sourcePath = "LargeWorkbook.xlsx";

        // Destination PDF file path
        string destPath = "LargeWorkbook.pdf";

        // Create a LightCells data handler (processes everything)
        var dataHandler = new SimpleLightCellsDataHandler();

        // Configure load options to use the LightCells data handler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = dataHandler;

        // Load the workbook using LightCells API
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Save the workbook directly to PDF
        workbook.Save(destPath, SaveFormat.Pdf);

        Console.WriteLine("Workbook successfully converted to PDF.");
    }

    // Minimal implementation of LightCellsDataHandler that processes all sheets, rows, and cells
    private class SimpleLightCellsDataHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet) => true;          // Process every sheet
        public bool StartRow(int rowIndex) => true;              // Process every row
        public bool ProcessRow(Row row) => true;                 // No custom row processing
        public bool StartCell(int columnIndex) => true;          // Process every cell
        public bool ProcessCell(Cell cell) => true;              // No custom cell processing
    }
}