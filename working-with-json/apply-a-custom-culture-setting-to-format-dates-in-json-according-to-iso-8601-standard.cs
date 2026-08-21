// Title: Import and Export JSON with ISO‑8601 Dates Using Aspose.Cells for .NET (Invariant Culture)
// Description: Demonstrates how to set Workbook.Settings.CultureInfo to InvariantCulture, import JSON containing ISO‑8601 date strings with JsonLayoutOptions (DateFormat = "yyyy-MM-ddTHH:mm:ssZ"), verify or convert cells to DateTime, and export the worksheet back to JSON while preserving the ISO‑8601 format. The workbook can also be saved as an Excel file.
// Keywords: Aspose.Cells JSON import | ISO 8601 date format | InvariantCulture workbook | JsonLayoutOptions DateFormat | JsonUtility ImportData | Export JSON Aspose.Cells | C# Excel to JSON | culture‑independent dates
// Common Searches: Aspose.Cells import JSON dates as DateTime | set workbook culture invariant Aspose.Cells .NET | export JSON with ISO 8601 format using Aspose.Cells | JsonLayoutOptions custom date format example | C# convert JSON string dates to Excel DateTime
// Developer Intent: The developer needs to read JSON data that contains ISO‑8601 date strings, store those dates as true DateTime values in an Excel workbook regardless of locale, and then write the worksheet back to JSON while keeping the ISO‑8601 representation.
// Use Cases: Load employee or transaction records from a locale‑neutral JSON feed, keep dates as DateTime in Excel for calculations, and generate a JSON export that downstream services can consume without additional parsing. | Create Excel reports from international JSON payloads where date separators must not depend on the server’s regional settings. | Validate imported date cells, re‑parse any string dates with DateTime.TryParseExact, and ensure consistent ISO‑8601 output for API integration.
// AI Prompts: Generate C# code that uses Aspose.Cells to import JSON with ISO‑8601 dates, convert them to DateTime cells, and export the sheet to JSON preserving the same format. | Explain the interaction between Workbook.Settings.CultureInfo and JsonLayoutOptions.DateFormat for culture‑independent date handling in Aspose.Cells. | Provide a step‑by‑step guide to detect string dates after import, convert them to DateTime, and re‑store them before exporting to JSON.

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonIsoDateDemo
{
    // Demonstrates how to set Workbook.Settings.CultureInfo to InvariantCulture, import JSON containing ISO‑8601 date strings with JsonLayoutOptions (DateFormat = "yyyy-MM-ddTHH:mm:ssZ"), verify or convert cells to DateTime, and export the worksheet back to JSON while preserving the ISO‑8601 format. The workbook can also be saved as an Excel file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Set the workbook culture to invariant to avoid locale‑specific date separators
                workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;

                // Prepare JSON data containing date strings
                string jsonData = @"[
                    { ""Id"": 1, ""Name"": ""Alice"", ""JoinDate"": ""2023-05-15T08:30:00Z"" },
                    { ""Id"": 2, ""Name"": ""Bob"",   ""JoinDate"": ""2023-06-20T14:45:00Z"" }
                ]";

                // Configure layout options:
                // - Convert strings that match the date format to actual DateTime values
                // - Use an ISO‑8601 date format when writing dates back to JSON
                JsonLayoutOptions layoutOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true,
                    ConvertNumericOrDate = true,
                    DateFormat = "yyyy-MM-ddTHH:mm:ssZ"   // ISO 8601
                };

                // Import the JSON data into the first worksheet starting at cell A1
                Worksheet sheet = workbook.Worksheets[0];
                JsonUtility.ImportData(jsonData, sheet.Cells, 0, 0, layoutOptions);

                // Verify that the dates are stored as DateTime values; if not, attempt conversion
                var cell = sheet.Cells["C2"];
                DateTime joinDate;

                if (cell.Value is DateTime dt)
                {
                    joinDate = dt;
                }
                else if (cell.Value is string s && DateTime.TryParseExact(
                             s,
                             "yyyy-MM-ddTHH:mm:ssZ",
                             CultureInfo.InvariantCulture,
                             DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                             out dt))
                {
                    // Store the parsed DateTime back into the cell
                    cell.PutValue(dt);
                    joinDate = dt;
                }
                else
                {
                    // Fallback: treat as MinValue to avoid exception
                    joinDate = DateTime.MinValue;
                }

                Console.WriteLine("Cell C2 type: " + cell.Type);
                Console.WriteLine("Cell C2 value (ISO 8601): " + joinDate.ToString("yyyy-MM-ddTHH:mm:ssZ"));

                // Save the workbook to an Excel file (optional)
                workbook.Save("Output.xlsx");

                // Export the worksheet back to JSON using the same ISO‑8601 format
                JsonSaveOptions saveOptions = new JsonSaveOptions
                {
                    Indent = "    ", // pretty‑print with 4 spaces
                    ExportArea = new CellArea
                    {
                        StartRow = 0,
                        EndRow = sheet.Cells.MaxDataRow,
                        StartColumn = 0,
                        EndColumn = sheet.Cells.MaxDataColumn
                    },
                    HasHeaderRow = true
                };

                // The DateFormat set in JsonLayoutOptions is also respected during export,
                // so dates will appear in ISO 8601 format.
                workbook.Save("ExportedData.json", saveOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
