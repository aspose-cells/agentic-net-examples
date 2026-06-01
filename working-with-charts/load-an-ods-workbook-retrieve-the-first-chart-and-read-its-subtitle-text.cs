using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the ODS workbook from file
        Workbook workbook = new Workbook("input.ods");

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Check that the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Retrieve the first chart in the collection
            Chart chart = worksheet.Charts[0];

            // Get the subtitle object (Title type) and read its text
            string subtitleText = chart.SubTitle.Text;

            // Output the subtitle text
            Console.WriteLine("Chart subtitle: " + subtitleText);
        }
        else
        {
            Console.WriteLine("No charts found in the workbook.");
        }
    }
}