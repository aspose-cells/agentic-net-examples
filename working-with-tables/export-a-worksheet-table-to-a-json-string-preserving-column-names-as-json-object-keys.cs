using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate the worksheet with a header row and some data rows
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");
                cells["C1"].PutValue("City");

                cells["A2"].PutValue("John");
                cells["B2"].PutValue(30);
                cells["C2"].PutValue("New York");

                cells["A3"].PutValue("Alice");
                cells["B3"].PutValue(25);
                cells["C3"].PutValue("London");

                cells["A4"].PutValue("Bob");
                cells["B4"].PutValue(28);
                cells["C4"].PutValue("Paris");

                // Define the range that includes the header and all data rows
                // start row = 0, start column = 0, total rows = 4, total columns = 3
                AsposeRange exportRange = cells.CreateRange(0, 0, 4, 3);

                // Configure JSON export options
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,          // first row contains column names
                    ExportEmptyCells = true,      // export empty cells as null
                    Indent = "    "               // pretty‑print with 4 spaces
                };

                // Export the defined range to a JSON string
                string jsonResult = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

                // Output the JSON string to the console
                Console.WriteLine("Exported JSON:");
                Console.WriteLine(jsonResult);
            }
            catch (Exception ex)
            {
                // Log the exception details
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}