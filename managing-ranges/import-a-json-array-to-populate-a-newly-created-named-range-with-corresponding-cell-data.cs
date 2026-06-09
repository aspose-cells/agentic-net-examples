using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonImportExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Sample JSON array (array of objects). Each object will become a row.
                string json = @"[
                    { ""Name"": ""John"", ""Age"": 30, ""City"": ""New York"" },
                    { ""Name"": ""Alice"", ""Age"": 25, ""City"": ""London"" },
                    { ""Name"": ""Bob"", ""Age"": 28, ""City"": ""Paris"" }
                ]";

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Configure JSON import options: treat the array as a table (adds header row).
                JsonLayoutOptions importOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true   // First row will contain property names.
                };

                // Import JSON data starting at cell A1 (row 0, column 0).
                JsonUtility.ImportData(json, cells, 0, 0, importOptions);

                // Determine the size of the imported data.
                // Header row + number of objects = 1 + 3 = 4 rows, 3 columns (Name, Age, City).
                int totalRows = 4;   // Adjust if your JSON size changes.
                int totalColumns = 3;

                // Create a named range that covers the imported data.
                Aspose.Cells.Range namedRange = cells.CreateRange(0, 0, totalRows, totalColumns);
                namedRange.Name = "PeopleData";

                // Optional: demonstrate that the named range works by using it in a formula.
                cells["E1"].Formula = "=SUM(PeopleData[Age])"; // Sum of the Age column.

                // Save the workbook.
                workbook.Save("PeopleData.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}