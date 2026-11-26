// Title: How to add a custom tooltip to a TextBox shape (InfoBox) in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set AlternativeText for a TextBox shape to display a tooltip in Excel | Create an InfoBox shape with hover help text using the Aspose.Cells .NET API | Programmatically add a tooltip to an Excel textbox shape with Aspose.Cells | Set a shape tooltip in a generated XLSX file using Aspose.Cells for .NET | Customize TextBox shape tooltip and styling with Aspose.Cells C#
// Tags: Aspose.Cells shape AlternativeText property | C# add tooltip to Excel TextBox shape | Aspose.Cells TextBox shape formatting | generate XLSX with shape tooltip using Aspose.Cells | Excel shape hover help text Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTooltipDemo
{
    // The example creates a new workbook, inserts a TextBox shape named 'InfoBox' on the first worksheet, assigns custom help text to its AlternativeText property to act as a tooltip, applies background and border styling, and saves the file as InfoBoxWithTooltip.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // If you have a license file, uncomment the following lines and set the correct path.
                // var license = new Aspose.Cells.License();
                // license.SetLicense("Aspose.Cells.lic");

                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a TextBox shape that will act as the InfoBox
                // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixels), width (pixels)
                TextBox infoBox = worksheet.Shapes.AddTextBox(2, 0, 2, 0, 120, 80);
                infoBox.Name = "InfoBox";
                infoBox.Text = "InfoBox";

                // Set the tooltip (alternative text) that appears when the user hovers over the shape
                // The AlternativeText property is used by Excel to display a tooltip.
                infoBox.AlternativeText = "Custom help text: This box provides additional information about the data.";

                // Optionally, you can style the TextBox (background color, border, etc.)
                infoBox.FillFormat.ForeColor = System.Drawing.Color.LightYellow;
                infoBox.LineFormat.Weight = 1.0;
                infoBox.LineFormat.ForeColor = System.Drawing.Color.DarkGray;

                // Save the workbook to a file
                string outputPath = "InfoBoxWithTooltip.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Open the file in Excel, hover over the shape named 'InfoBox' to see the tooltip.");
            }
            catch (Exception ex)
            {
                // Log the exception details to the console
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
