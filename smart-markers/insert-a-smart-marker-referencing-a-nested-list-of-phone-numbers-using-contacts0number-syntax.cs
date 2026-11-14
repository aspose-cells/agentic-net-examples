// Title: Insert a smart marker for the first contact's phone number (contacts[0].number) into an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# insert smart marker for nested list element contacts[0].number | How to reference a phone number from a contacts array in an Excel smart marker with Aspose.Cells | C# example of using ${contacts[0].number} smart marker in Aspose.Cells workbook | Create Excel file with smart marker that pulls first contact's number using Aspose.Cells .NET | Smart marker syntax for accessing first item in a list in Aspose.Cells C#
// Tags: insert smart marker Aspose.Cells C# | nested list reference smart marker Excel | contacts array phone number marker Aspose.Cells | Aspose.Cells workbook save with smart marker | Excel smart marker syntax contacts[0].number

using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, places the smart marker "${contacts[0].number}" in cell A1 to pull the first contact's phone number from a nested list, and saves the workbook as SmartMarkerExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a smart marker that references the first contact's number (nested list)
        // Syntax: contacts[0].number
        sheet.Cells["A1"].PutValue("${contacts[0].number}");

        // Save the workbook
        workbook.Save("SmartMarkerExample.xlsx");
    }
}
