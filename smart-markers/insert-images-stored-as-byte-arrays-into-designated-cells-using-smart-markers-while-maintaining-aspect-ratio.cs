using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertImagesWithSmartMarkers
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the template workbook containing the smart marker "&=Photo"
            Workbook workbook = new Workbook(templatePath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a DataTable that holds image byte arrays
            DataTable table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Photo", typeof(byte[]));

            // Load images from files into byte arrays (check existence first)
            const string imgPath1 = "image1.png";
            const string imgPath2 = "image2.jpg";

            if (!File.Exists(imgPath1) || !File.Exists(imgPath2))
            {
                Console.WriteLine("One or more image files were not found.");
                return;
            }

            byte[] imageBytes1 = File.ReadAllBytes(imgPath1);
            byte[] imageBytes2 = File.ReadAllBytes(imgPath2);

            // Add rows to the table
            table.Rows.Add("First Image", imageBytes1);
            table.Rows.Add("Second Image", imageBytes2);

            // Put the table into a DataSet – the smart marker processor works with DataSet objects
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(table);

            // Process the smart markers using WorkbookDesigner (the recommended API)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Ensure each inserted picture is placed inside the cell.
            // Aspect ratio locking is handled automatically by Aspose.Cells when inserting via smart markers.
            foreach (Picture picture in worksheet.Pictures)
            {
                picture.IsPlacedInCell = true; // embed picture in cell
            }

            // Save the workbook with the embedded images
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}