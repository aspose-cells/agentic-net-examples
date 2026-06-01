using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Markup;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a smart tag to cell A1 (row 0, column 0)
            SmartTagSetting smartTagSetting = worksheet.SmartTagSetting;
            smartTagSetting.Add(0, 0); // creates a SmartTagCollection for the cell

            // Retrieve the SmartTagCollection for the cell and add a "date" smart tag
            SmartTagCollection smartTagCollection = smartTagSetting[0, 0];
            smartTagCollection.Add("urn:schemas-microsoft-com:office:smarttags", "date");

            // Insert the TODAY() formula into the same cell
            Cell cell = worksheet.Cells[0, 0];
            cell.Formula = "=TODAY()";

            // Save the workbook
            string outputPath = "SmartTagWithToday.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}