using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    class AddHyperlinksToNamedRange
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Item1");
            worksheet.Cells["A2"].PutValue("Item2");
            worksheet.Cells["A3"].PutValue("Item3");

            // Define a named range "MyRange" that covers cells A1:A3
            int firstRow = 0;          // Row index for A1
            int firstColumn = 0;       // Column index for A1
            int totalRows = 3;         // A1, A2, A3
            int totalColumns = 1;      // Single column
            worksheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns).Name = "MyRange";

            // Retrieve the range by its name using the workbook's Worksheets collection
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("MyRange");

            // Iterate through each cell in the named range and add a hyperlink
            for (int r = 0; r < namedRange.RowCount; r++)
            {
                for (int c = 0; c < namedRange.ColumnCount; c++)
                {
                    // Get the address of the current cell (e.g., "A1")
                    string cellAddress = namedRange[r, c].Name;

                    // Construct the documentation URL for this cell
                    string docUrl = $"https://docs.mycompany.com/{cellAddress}";

                    // Add a hyperlink to the cell
                    int hyperlinkIndex = worksheet.Hyperlinks.Add(cellAddress, 1, 1, docUrl);

                    // Set a friendly display text for the hyperlink
                    worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = $"Doc {cellAddress}";
                }
            }

            // Save the workbook with the added hyperlinks
            string outputPath = "NamedRangeHyperlinks.xlsx";
            workbook.Save(outputPath);
        }
    }
}