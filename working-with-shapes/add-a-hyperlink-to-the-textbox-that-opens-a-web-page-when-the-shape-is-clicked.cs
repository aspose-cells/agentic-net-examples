// Title: Add a clickable hyperlink to a textbox shape in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a textbox shape on a worksheet and assigns a Hyperlink.Address pointing to a given URL with Aspose.Cells. | Write a C# program that adds a textbox to an Excel sheet, sets its displayed text, and makes it open a web page when clicked using Aspose.Cells. | Update an existing Aspose.Cells workbook to attach a hyperlink to a shape so the shape behaves as a clickable link.
// Common Searches: Aspose.Cells C# add hyperlink to textbox shape in Excel | set hyperlink address on a shape using Aspose.Cells .NET | create clickable textbox in Excel file with Aspose.Cells library | how to assign web URL to a textbox shape in Aspose.Cells C# example
// Tags: Aspose.Cells add textbox shape | Aspose.Cells set shape hyperlink | C# Aspose.Cells hyperlink textbox | Excel workbook shape hyperlink .NET | Aspose.Cells Hyperlink.Address property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // C# example that creates a new workbook, adds a textbox shape to the first worksheet, sets its text, assigns a Hyperlink.Address to https://www.example.com, and saves the file as HyperlinkTextbox.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, height, width
                Shape textBox = worksheet.Shapes.AddTextBox(2, 2, 5, 5, 100, 200);

                // Set the displayed text of the textbox
                textBox.Text = "Click here to visit Example.com";

                // Assign a hyperlink to the textbox so it opens a web page when clicked
                // The Hyperlink property is read‑only; modify its Address instead
                textBox.Hyperlink.Address = "https://www.example.com";

                // Define output file path
                string outputPath = "HyperlinkTextbox.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
