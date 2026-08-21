// Title: C# – Apply Light‑Blue Solid Fill and Bold Font Style to a DataSet Range with Aspose.Cells
// Description: Shows how to create a Style in Aspose.Cells for .NET, set a solid light‑blue background and bold font, apply it to a range (A1:C5) representing a data set, and save the workbook.
// Keywords: Aspose.Cells C# style solid fill | light blue background Aspose.Cells | bold font style range | apply style to range Aspose.Cells | create reusable style .NET | Excel formatting Aspose.Cells
// Common Searches: Aspose.Cells set solid fill color for cells | C# apply bold font to a range in Excel using Aspose | How to style a dataset range with Aspose.Cells | Create and reuse Style object Aspose.Cells .NET | Set background color of a range Aspose.Cells
// Developer Intent: The developer wants to format a specific cell range with a solid light‑blue background and bold text using Aspose.Cells for .NET.
// Use Cases: Generate a report where header rows are highlighted with a light‑blue background and bold font. | Standardize the appearance of imported data tables before exporting them to Excel. | Create a reusable Style object to format multiple sections across several worksheets.
// AI Prompts: Provide C# code that creates a reusable Style with a solid light‑blue fill and bold font, then applies it to multiple ranges in an Aspose.Cells workbook. | Show how to change the foreground color of an existing Style to light blue without recreating the Style object in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Shows how to create a Style in Aspose.Cells for .NET, set a solid light‑blue background and bold font, apply it to a range (A1:C5) representing a data set, and save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style using the Workbook.CreateStyle method (rule: createstyle)
            Style style = workbook.CreateStyle();
            // Set solid fill with light blue foreground color
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.LightBlue;
            // Make the font bold
            style.Font.IsBold = true;

            // Define the range that represents the DataSet (example: A1:C5)
            AsposeRange dataSetRange = worksheet.Cells.CreateRange("A1", "C5");

            // Apply the style to the entire range (rule: setstyle)
            dataSetRange.SetStyle(style);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("StyledDataSet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
