// Title: Export an Aspose.Cells workbook to CSV with ISO 8601 date formatting in C#
// AI Prompts: Generate C# code that creates a custom style using the pattern yyyy‑MM‑ddTHH:mm:ss, applies it to DateTime cells, and saves the workbook as CSV with Aspose.Cells. | Show how to configure Aspose.Cells so that date values are written in ISO 8601 format during CSV export. | Provide a step‑by‑step example that inserts DateTime values, sets an ISO 8601 style, and calls Workbook.Save with SaveFormat.Csv.
// Common Searches: how to export Excel to CSV with ISO 8601 dates using Aspose.Cells in C# | Aspose.Cells C# custom date format for CSV export | C# save workbook as CSV preserving date format yyyy-MM-ddTHH:mm:ss | set ISO 8601 date style in Aspose.Cells before CSV conversion | Aspose.Cells CSV output date formatting options .NET
// Tags: Aspose.Cells CSV export with custom date style | ISO 8601 date formatting in Aspose.Cells | C# workbook.Save SaveFormat.Csv custom date format | apply custom style to DateTime cells Aspose.Cells | export worksheet to CSV preserving date format

using System;
using Aspose.Cells;

// The example creates a workbook, writes DateTime values, defines a custom style with the ISO 8601 pattern yyyy‑MM‑ddTHH:mm:ss, applies the style to the date cells, and saves the workbook as a CSV file where the dates appear in standardized ISO 8601 format.
class ExportWorkbookToCsvIso
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with DateTime values
        worksheet.Cells["A1"].PutValue(DateTime.Now);
        worksheet.Cells["A2"].PutValue(new DateTime(2023, 5, 15, 13, 45, 30));

        // Define ISO 8601 date format (e.g., 2023-05-15T13:45:30)
        Style isoStyle = workbook.CreateStyle();
        isoStyle.Custom = "yyyy-MM-ddTHH:mm:ss";

        // Apply the ISO format style to the date cells
        worksheet.Cells["A1"].SetStyle(isoStyle);
        worksheet.Cells["A2"].SetStyle(isoStyle);

        // Save the workbook as CSV; dates will be output using the custom ISO format
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
