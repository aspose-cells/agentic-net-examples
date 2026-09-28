// Title: Insert the current date with a smart tag using the TODAY() formula in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a smart tag to cell A1, assigns the =TODAY() formula, and saves the file as an .xlsx. | Generate a snippet showing how to configure SmartTagSetting and SmartTagCollection to display today's date via a smart marker in Aspose.Cells. | Provide an example that applies a "date" smart tag with the TODAY() function and exports the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add smart tag that shows today's date | how to use TODAY() formula with smart markers in Aspose.Cells .NET | create Excel file with dynamic date using smart tag in C# Aspose.Cells | smart tag insert current date Aspose.Cells example
// Tags: smart tag add date Aspose.Cells C# | set cell formula TODAY() Aspose.Cells | configure SmartTagSetting Aspose.Cells | dynamic date insertion Excel Aspose.Cells | save workbook with smart marker .xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Markup;

// Demonstrates creating a workbook, adding a 'date' smart tag to cell A1, setting its formula to =TODAY() for dynamic current‑date display, and saving the result as SmartTagWithToday.xlsx.
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
            smartTagSetting.Add(0, 0); // Add(int row, int column)

            // Retrieve the SmartTagCollection for the cell and add a "date" smart tag
            SmartTagCollection smartTagCollection = smartTagSetting[0, 0];
            smartTagCollection.Add("urn:schemas-microsoft-com:office:smarttags", "date"); // Add(string uri, string name)

            // Set the cell formula to TODAY() so it displays the current date
            worksheet.Cells[0, 0].Formula = "=TODAY()";

            // Save the workbook
            workbook.Save("SmartTagWithToday.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
