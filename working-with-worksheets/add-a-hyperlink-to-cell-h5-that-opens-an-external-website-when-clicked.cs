using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a hyperlink to cell H5 that points to an external website
            // Parameters: cell name, number of rows, number of columns, hyperlink address
            worksheet.Hyperlinks.Add("H5", 1, 1, "https://www.example.com");

            // Optionally set the display text for the hyperlink
            worksheet.Hyperlinks[0].TextToDisplay = "Visit Example Site";

            // Save the workbook to a file
            workbook.Save("HyperlinkH5.xlsx");
        }
    }
}