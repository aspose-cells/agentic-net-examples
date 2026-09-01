// Title: Convert an Excel workbook to CSV in C# with Aspose.Cells while preserving blank rows and column positions
// AI Prompts: Generate C# code that loads an .xlsx file and saves it as a CSV using Aspose.Cells with TrimLeadingBlankRowAndColumn set to false and KeepSeparatorsForBlankRow enabled. | Show how to configure TxtSaveOptions for CSV export to retain the original column layout and include separators for empty rows in Aspose.Cells. | Provide a complete example that converts a workbook to CSV without removing leading blank rows or columns, preserving the exact spreadsheet structure.
// Common Searches: Aspose.Cells C# export to CSV keep blank rows and columns | How to prevent trimming of leading blank rows when saving Excel as CSV with Aspose.Cells | TxtSaveOptions KeepSeparatorsForBlankRow true example in .NET | Preserve column positions during CSV conversion using Aspose.Cells | Save Excel workbook as CSV without losing empty rows Aspose.Cells .NET
// Tags: Aspose.Cells CSV export preserving blanks | TxtSaveOptions TrimLeadingBlankRowAndColumn false | KeepSeparatorsForBlankRow Aspose.Cells | C# Excel to CSV conversion with layout retention | SaveFormat.Csv Aspose.Cells configuration

using System;
using Aspose.Cells;

// Loads an Excel workbook, configures TxtSaveOptions with TrimLeadingBlankRowAndColumn = false and KeepSeparatorsForBlankRow = true, then saves the file as a CSV, preserving the original column layout and blank rows.
class WorkbookToCsv
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure CSV save options:
        // - Do not trim leading blank rows/columns so the original column layout is preserved.
        // - Keep separators for completely blank rows to maintain row count.
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            TrimLeadingBlankRowAndColumn = false,
            KeepSeparatorsForBlankRow = true
        };

        // Save the workbook as CSV using the configured options
        workbook.Save("output.csv", csvOptions);
    }
}
