using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

class ExportWorksheetToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.9);

            // Apply a header style
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.White;
            headerStyle.ForegroundColor = System.Drawing.Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Apply a numeric style to price cells
            Style priceStyle = workbook.CreateStyle();
            priceStyle.Number = 2; // two decimal places
            sheet.Cells["B2"].SetStyle(priceStyle);
            sheet.Cells["B3"].SetStyle(priceStyle);

            // Define the range to export (A1:B3)
            AsposeRange exportRange = sheet.Cells.CreateRange("A1:B3");

            // Configure JSON export options for custom serialization
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,      // export all values as strings
                ExportEmptyCells = true,    // include empty cells as null
                ExportStylePool = false,    // export style for each cell individually
                HasHeaderRow = true,        // first row contains headers
                Indent = "  ",              // pretty‑print JSON with two‑space indentation
                ToExcelStruct = true        // include Excel structure information
            };

            // Convert the range to JSON using the configured options
            string json = exportRange.ToJson(jsonOptions);

            // Output the JSON to the console
            Console.WriteLine(json);

            // Save the JSON string to a file
            string outputPath = "WorksheetExport.json";
            File.WriteAllText(outputPath, json);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}