using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;   // Needed for ListObject

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // OPTIONAL: add a slicer so the collection is not empty (remove if not needed)
            // Create sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);

            // Add a table (ListObject) covering the sample data
            int tableIdx = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = worksheet.ListObjects[tableIdx];

            // Add a slicer for the first column of the table
            worksheet.Slicers.Add(table, 0, "D1");

            // Retrieve the slicer collection from the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Iterate through each slicer and log its name
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];
                Console.WriteLine("Slicer Name: " + slicer.Name);
            }

            // Define output file path
            string outputPath = "SlicersLogDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}