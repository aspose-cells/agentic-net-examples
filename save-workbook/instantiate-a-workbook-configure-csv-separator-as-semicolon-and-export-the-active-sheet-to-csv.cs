// Title: Save the active worksheet as a CSV file with a semicolon delimiter using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save active worksheet to CSV with semicolon separator | How to set custom field delimiter when exporting CSV with Aspose.Cells .NET | C# code to export first worksheet to CSV using semicolon as delimiter in Aspose.Cells | Configure TxtSaveOptions to use ';' for CSV output in Aspose.Cells | Export workbook's active sheet to CSV file with custom separator in .NET
// Tags: TxtSaveOptions CSV delimiter semicolon | configure CSV field separator Aspose.Cells | export worksheet as CSV using Aspose.Cells | set custom CSV separator in .NET workbook save | Aspose.Cells save format CSV with custom delimiter

using Aspose.Cells;
using System;

// The example creates an empty workbook, sets TxtSaveOptions to CSV format with a semicolon as the field separator, and saves the active (first) worksheet to a file named ActiveSheetExport.csv.
class Program
{
    static void Main()
    {
        // Instantiate a new workbook (empty workbook with a default worksheet)
        Workbook workbook = new Workbook();

        // Configure CSV save options to use semicolon as the field separator
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
        csvOptions.Separator = ';';

        // Export the active sheet (the first worksheet) to a CSV file
        workbook.Save("ActiveSheetExport.csv", csvOptions);
    }
}
