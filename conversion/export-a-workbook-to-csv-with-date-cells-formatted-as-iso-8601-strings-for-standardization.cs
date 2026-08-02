// Title: C# – Export Aspose.Cells Workbook to CSV with ISO 8601 Date Formatting
// Description: Creates a workbook, inserts DateTime values, applies a custom "yyyy-MM-ddTHH:mm:ss" style, and saves the file as CSV so all date cells are written in ISO 8601 format.
// Keywords: Aspose.Cells CSV export C# | ISO 8601 date format Aspose | C# custom date style CSV | Aspose.Cells export to CSV with dates | CSV date formatting Aspose.Cells | Save workbook as CSV ISO date
// Common Searches: Aspose.Cells export CSV with ISO 8601 dates C# | How to format dates as ISO 8601 when saving CSV using Aspose.Cells | C# Aspose.Cells custom date style before CSV export | Save workbook to CSV preserving date format Aspose | ISO 8601 timestamp CSV Aspose.Cells example
// Developer Intent: Generate a CSV file from an Aspose.Cells workbook where every date cell is rendered as an ISO 8601 timestamp.
// Use Cases: Standardizing timestamps for data pipelines that require ISO 8601 strings. | Producing CSV reports that can be parsed reliably by external systems or APIs. | Creating audit logs in CSV with consistent, sortable date formats.
// AI Prompts: Show C# code that applies an ISO 8601 custom date format to cells and saves the workbook as CSV using Aspose.Cells. | Give an example of exporting an Aspose.Cells workbook to CSV while ensuring all dates appear as "yyyy-MM-ddTHH:mm:ss".

using System;
using Aspose.Cells;

// Creates a workbook, inserts DateTime values, applies a custom "yyyy-MM-ddTHH:mm:ss" style, and saves the file as CSV so all date cells are written in ISO 8601 format.
class ExportWorkbookToCsvWithIsoDate
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert sample date values
        cells["A1"].PutValue(DateTime.Now);                         // Current date & time
        cells["A2"].PutValue(new DateTime(2023, 5, 15, 13, 45, 30)); // Fixed date & time

        // Define ISO 8601 format (e.g., 2023-05-15T13:45:30)
        string iso8601Format = "yyyy-MM-ddTHH:mm:ss";

        // Create a style with the custom date format
        Style isoStyle = workbook.CreateStyle();
        isoStyle.Custom = iso8601Format;

        // Apply the style to the cells containing dates
        cells["A1"].SetStyle(isoStyle);
        cells["A2"].SetStyle(isoStyle);

        // Save the workbook as CSV; the dates will be written using the ISO 8601 format
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
