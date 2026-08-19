// Title: Insert Current Date with a Smart Tag Using TODAY() in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a smart tag to cell A1, assign the "date" smart tag type, set the cell formula to =TODAY() so the date updates automatically, and save the file as SmartTagWithToday.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | smart tag | TODAY() function | C# | .NET | insert current date | Excel automation | smart markers | dynamic date formula | worksheet smart tag
// Common Searches: Aspose.Cells add smart tag with TODAY() | C# set cell formula to TODAY() in Aspose.Cells | how to use smart tags for dates in Excel with Aspose | dynamic date smart marker Aspose.Cells .NET | create workbook with smart tag that shows today's date
// Developer Intent: Add a smart tag that automatically displays the current date by applying the TODAY() formula to a cell.
// Use Cases: Generate report templates where the header shows the opening date that refreshes on each view. | Automate invoice sheets that stamp the creation date without manual entry. | Build export routines that add a summary cell reflecting the current processing date.
// AI Prompts: Write C# code using Aspose.Cells to add a smart tag to cell B2 that inserts the current date with the TODAY() function. | Explain how smart tags and formulas can be combined in Aspose.Cells to display dynamic values such as the current date. | Provide example with robust error handling for adding a date smart tag and setting the TODAY() formula in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Markup;

// Demonstrates how to create a workbook, add a smart tag to cell A1, assign the "date" smart tag type, set the cell formula to =TODAY() so the date updates automatically, and save the file as SmartTagWithToday.xlsx using Aspose.Cells for .NET.
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
            smartTagSetting.Add(0, 0); // index not needed further

            // Retrieve the SmartTagCollection for the cell and add a "date" smart tag
            SmartTagCollection smartTagCollection = smartTagSetting[0, 0];
            smartTagCollection.Add("urn:schemas-microsoft-com:office:smarttags", "date");

            // Insert the current date using the TODAY() function
            worksheet.Cells["A1"].Formula = "=TODAY()";

            // Save the workbook
            workbook.Save("SmartTagWithToday.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
