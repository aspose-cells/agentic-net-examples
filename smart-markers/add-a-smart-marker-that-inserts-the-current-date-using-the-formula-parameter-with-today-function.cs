// Title: C# Example: Add a Smart Tag with TODAY() Formula to Insert Current Date using Aspose.Cells .NET
// Description: This Aspose.Cells for .NET snippet creates a workbook, accesses the first worksheet, configures a SmartTagSetting for cell A1, adds a "date" smart tag, assigns the TODAY() formula so the cell always shows the current date, and saves the file as SmartTagWithToday.xlsx.
// Keywords: Aspose.Cells | C# | .NET | SmartTag | SmartTagSetting | SmartTagCollection | TODAY() function | current date formula | Excel automation | date smart tag | sample code | GitHub example
// Common Searches: Aspose.Cells add smart tag TODAY() | C# set smart tag formula to TODAY | Insert current date with smart tag Aspose.Cells | SmartTagSetting example .NET | How to use TODAY() in Aspose.Cells smart tag
// Developer Intent: Insert a smart tag that automatically displays today’s date by using the TODAY() formula in an Aspose.Cells workbook.
// Use Cases: Create a daily report template where the header cell always shows the current date. | Generate spreadsheets programmatically with a date placeholder that updates on each open. | Build a reusable Excel workbook that includes a smart tag for dynamic date insertion alongside other data fields.
// AI Prompts: Generate C# code that adds a smart tag to cell B2 and sets its formula to =TODAY() using Aspose.Cells. | Show how to add multiple smart tags with different formulas (e.g., TODAY(), NOW()) in a single Aspose.Cells workbook. | Explain how to locate an existing smart tag in a worksheet and modify its formula to use a different date function.

using System;
using Aspose.Cells;
using Aspose.Cells.Markup;

// This Aspose.Cells for .NET snippet creates a workbook, accesses the first worksheet, configures a SmartTagSetting for cell A1, adds a "date" smart tag, assigns the TODAY() formula so the cell always shows the current date, and saves the file as SmartTagWithToday.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a SmartTagCollection to cell A1 (row 0, column 0)
            SmartTagSetting smartTagSetting = worksheet.SmartTagSetting;
            smartTagSetting.Add(0, 0); // creates the collection for A1

            // Retrieve the collection and add a "date" smart tag
            SmartTagCollection smartTagCollection = smartTagSetting[0, 0];
            smartTagCollection.Add("urn:schemas-microsoft-com:office:smarttags", "date");

            // Insert the current date using the TODAY() formula
            worksheet.Cells["A1"].Formula = "=TODAY()";

            // Save the workbook
            workbook.Save("SmartTagWithToday.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
