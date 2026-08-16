// Title: Aspose.Cells .NET: Create A1:C10 named range and rename to ReportData (C#)
// Description: C# example that creates a new workbook, defines the range A1:C10, assigns it the name "DataSet", then renames it to "ReportData" using the Name.Text property, and saves the file as NamedRangeDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | named range | CreateRange | Workbook.Names | Name.Text | rename named range | Excel automation | ReportData | A1:C10
// Common Searches: Aspose.Cells create named range C# | rename named range Aspose.Cells | Name.Text property Aspose.Cells example | define range A1:C10 Aspose.Cells | how to change named range name in .NET workbook
// Developer Intent: Create a named range covering cells A1:C10 and then change its name to ReportData programmatically with Aspose.Cells for .NET.
// Use Cases: Build a template where the data block A1:C10 is referenced by a meaningful name that can be updated later. | Automate renaming of an existing named range to match a new reporting period or naming convention. | Prepare spreadsheets for formulas that rely on a named range, then adjust the name without altering cell references.
// AI Prompts: Generate C# code that defines a named range A1:C10 and renames it to ReportData using Aspose.Cells. | Explain how to retrieve a named range from Workbook.Names and modify its Text property in Aspose.Cells for .NET. | Compare Name.Text and Name.RefersTo when managing named ranges with Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that creates a new workbook, defines the range A1:C10, assigns it the name "DataSet", then renames it to "ReportData" using the Name.Text property, and saves the file as NamedRangeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Sheet1";

            // Define the range A1:C10 (use fully qualified Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:C10");

            // Assign the initial name "DataSet" to the range
            range.Name = "DataSet";

            // Retrieve the Name object from the workbook's Names collection
            Name namedRange = workbook.Worksheets.Names["DataSet"];

            // Rename the range to "ReportData" using the Text property
            namedRange.Text = "ReportData";

            // Save the workbook
            workbook.Save("NamedRangeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
