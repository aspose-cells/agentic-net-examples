using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a smart marker that will be replaced by an image.
            // The marker name "ImagePath" will be matched with the data source key.
            worksheet.Cells["B2"].PutValue("&=ImagePath");

            // Prepare the data source: the value is a file path to the image.
            string imagePath = "sample.jpg"; // replace with your actual image file path

            // Prevent FileNotFoundException for the image file
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Build a DataSet with a single table containing the image path.
            DataTable table = new DataTable("Data");
            table.Columns.Add("ImagePath", typeof(string));
            table.Rows.Add(imagePath);

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(table);

            // Process the smart markers using WorkbookDesigner (lifecycle rule: load)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the workbook (lifecycle rule: save)
            string outputPath = "SmartMarkerImage.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}