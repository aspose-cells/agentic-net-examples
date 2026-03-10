using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class LoadWorkbookForXmlExport
    {
        static void Main()
        {
            // Path to the source workbook (any supported format: .xlsx, .xls, .csv, .xml, etc.)
            string inputPath = "sourceWorkbook.xlsx";

            // Load the workbook using the lifecycle rule
            Workbook workbook = new Workbook(inputPath);

            // At this point the workbook is loaded and ready for further XML export operations.
            Console.WriteLine("Workbook loaded successfully. Ready for XML export.");
        }
    }
}