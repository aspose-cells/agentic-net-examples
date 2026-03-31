using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace OdsJsonToCsvDemo
{
    class Program
    {
        static void Main()
        {
            // Sample ODS JSON data (replace with actual JSON source)
            string odsJson = @"{
                ""Sheet1"": [
                    { ""Name"": ""John"", ""Age"": 30, ""City"": ""New York"" },
                    { ""Name"": ""Alice"", ""Age"": 25, ""City"": ""London"" },
                    { ""Name"": ""Bob"", ""Age"": 35, ""City"": ""Paris"" }
                ]
            }";

            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Configure JSON layout options:
            // - ArrayAsTable = true  => each array becomes a table with a header row
            // - ConvertNumericOrDate = true => numeric strings become numbers, dates become DateTime
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true,
                DateFormat = "yyyy-MM-dd",
                NumberFormat = "0.##"
            };

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            // Using the ImportData method as defined in the rule set
            JsonUtility.ImportData(odsJson, cells, 0, 0, layoutOptions);

            // Save the workbook as CSV.
            // Aspose.Cells automatically handles proper encoding (UTF-8 by default).
            // The CSV will contain the data imported from the JSON, preserving field mappings.
            workbook.Save("Output.csv", SaveFormat.Csv);

            Console.WriteLine("ODS JSON data has been successfully converted to CSV: Output.csv");
        }
    }
}