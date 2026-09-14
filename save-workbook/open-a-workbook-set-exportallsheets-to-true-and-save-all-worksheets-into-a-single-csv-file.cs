// Title: How to export every worksheet from an Excel workbook into a single CSV file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets TxtSaveOptions.ExportAllSheets to true, and writes all sheets to one CSV file. | Show how to configure Aspose.Cells TxtSaveOptions for CSV to combine multiple worksheets into a single output in a .NET application.
// Common Searches: Aspose.Cells C# export all workbook sheets to a single CSV | Save multiple Excel worksheets as one CSV using TxtSaveOptions | How to combine all sheets into one CSV file with Aspose.Cells .NET | ExportAllSheets property example for CSV output in Aspose.Cells
// Tags: Aspose.Cells export all sheets CSV | TxtSaveOptions ExportAllSheets .NET | C# combine workbook worksheets into single CSV | Save workbook as combined CSV using Aspose.Cells | Export multiple Excel sheets to one CSV file

using System;
using Aspose.Cells;

namespace AsposeCellsExportAllSheetsToCsv
{
    // Loads an existing .xlsx workbook, enables TxtSaveOptions.ExportAllSheets, and saves all worksheets together into a single CSV file with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create CSV save options and enable exporting all worksheets
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            csvOptions.ExportAllSheets = true;

            // Save all worksheets into a single CSV file
            workbook.Save("output_all_sheets.csv", csvOptions);
        }
    }
}
