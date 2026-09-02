// Title: Save an Aspose.Cells workbook as a UTF-8 CSV file with a Byte Order Mark (BOM) in C#
// AI Prompts: Write C# code that creates a workbook, populates cells, and uses TxtSaveOptions to export it to a UTF-8 CSV file that includes a BOM. | Show how to set TxtSaveOptions.Encoding to Encoding.UTF8 so Aspose.Cells writes a Byte Order Mark when saving a CSV file. | Provide a complete example that saves a workbook to "output_with_bom.csv" using the correct save options for BOM in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# export to CSV with UTF-8 BOM for Excel compatibility | How to force a Byte Order Mark when saving CSV using Aspose.Cells .NET | TxtSaveOptions SaveFormat.Csv encoding BOM example in C# | C# code to generate CSV from workbook with UTF-8 BOM using Aspose.Cells
// Tags: Aspose.Cells TxtSaveOptions CSV BOM | C# export workbook to UTF-8 CSV Aspose.Cells | CSV export with Byte Order Mark Aspose.Cells .NET | SaveFormat.Csv encoding UTF8 Aspose.Cells

using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvBomDemo
{
    // // Creates a workbook, fills sample data, configures TxtSaveOptions with Encoding = Encoding.UTF8 (which writes a Byte Order Mark), and saves the workbook as "output_with_bom.csv" in CSV format.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Configure CSV save options with UTF-8 encoding (includes BOM)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            csvOptions.Encoding = Encoding.UTF8; // ensures a Byte Order Mark is written

            // Save the workbook as CSV with the specified options
            workbook.Save("output_with_bom.csv", csvOptions);
        }
    }
}
