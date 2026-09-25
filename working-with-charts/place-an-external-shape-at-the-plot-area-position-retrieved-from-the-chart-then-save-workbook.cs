// Title: Insert an external PNG as a picture shape over a chart's plot area and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook, read the first chart's PlotArea X/Y/Width/Height, and add a picture shape from a PNG file stream that matches those dimensions with Aspose.Cells in C#. | Create a Shape via sheet.Shapes.AddPicture, set its Left, Top, Width, and Height to the chart's PlotArea values, then save the modified workbook. | Retrieve chart plot area coordinates using Aspose.Cells, align an external image shape to those coordinates, and persist the changes to a new XLSX file.
// Common Searches: Aspose.Cells C# add picture shape over chart plot area | how to get chart PlotArea coordinates with Aspose.Cells .NET | place external PNG on chart area using Aspose.Cells for .NET | save workbook after inserting image shape into chart plot area Aspose.Cells | C# align picture shape to chart PlotArea dimensions Aspose.Cells
// Tags: add picture shape from file stream Aspose.Cells | chart plot area coordinates Aspose.Cells C# | align external PNG with chart plot area Aspose.Cells | set shape dimensions using chart.PlotArea Aspose.Cells | save workbook after modifying shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads an existing XLSX file, obtains the first chart's plot area position and size, inserts an external PNG as a picture shape sized to exactly cover that plot area, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string workbookPath = "input.xlsx";
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the sheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Retrieve the plot area position and size (X/Y are the top‑left coordinates)
            double plotX = chart.PlotArea.X;
            double plotY = chart.PlotArea.Y;
            double plotWidth = chart.PlotArea.Width;
            double plotHeight = chart.PlotArea.Height;

            // Path to the external image
            string imagePath = "externalImage.png";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add a picture shape (initially zero size) using a file stream
            using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                Shape pictureShape = sheet.Shapes.AddPicture(0, 0, 0, 0, imgStream);

                // Position the picture shape to match the chart's plot area
                pictureShape.Left = (int)plotX;
                pictureShape.Top = (int)plotY;
                pictureShape.Width = (int)plotWidth;
                pictureShape.Height = (int)plotHeight;
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
