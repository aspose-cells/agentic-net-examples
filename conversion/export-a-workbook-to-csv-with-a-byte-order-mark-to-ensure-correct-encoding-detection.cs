// Title: How to export an Aspose.Cells workbook to CSV with a UTF-8 Byte Order Mark using C#
// Common Searches: Aspose.Cells C# save workbook as CSV with UTF-8 BOM | C# TxtSaveOptions CSV encoding include byte order mark | Generate CSV with UTF-8 BOM using Aspose.Cells library | Set CSV separator and BOM in Aspose.Cells .NET | Export Excel to CSV with UTF-8 BOM in .NET Core
// Tags: Aspose.Cells CSV export with UTF-8 BOM | TxtSaveOptions UTF-8 encoding Aspose.Cells | C# workbook to CSV with byte order mark | CSV separator configuration Aspose.Cells | SaveFormat.CSV options Aspose.Cells | UTF8Encoding true BOM Aspose.Cells | Export Excel to CSV .NET

using System.Text;
using Aspose.Cells;

// The example creates a Workbook, adds sample data, configures TxtSaveOptions for CSV format with UTF-8 encoding that emits a Byte Order Mark, sets the comma separator, and saves the file as 'output.csv' using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Populate the first worksheet with sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(30);

        // Configure CSV save options to include a UTF-8 BOM
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
        csvOptions.Encoding = new UTF8Encoding(true); // true => emit BOM
        csvOptions.Separator = ','; // standard comma separator

        // Export the workbook to CSV with the specified options
        workbook.Save("output.csv", csvOptions);
    }
}
