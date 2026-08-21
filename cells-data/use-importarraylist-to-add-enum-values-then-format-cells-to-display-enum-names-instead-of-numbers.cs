// Title: C# – Import enum values with ImportArrayList and display enum names via custom number format in Aspose.Cells
// Description: Shows how to create an ArrayList of enum integers, import it into a worksheet using ImportArrayList, define a Style with a custom number format that maps 0‑2 to "Unknown", "Active", "Inactive", apply the style to the range, and save the workbook.
// Keywords: Aspose.Cells | ImportArrayList | enum | custom number format | C# | .NET | Excel export | display enum names | style formatting
// Common Searches: Aspose.Cells ImportArrayList enum | display enum names in Excel C# | custom number format for enum values Aspose.Cells | map numeric codes to text in Excel using Aspose | C# Aspose.Cells format cells as text
// Developer Intent: Load enum numeric values into a spreadsheet and format the cells so the corresponding enum names appear instead of numbers.
// Use Cases: Export a list of status codes to Excel with readable labels. | Create reports where enum columns are shown as text rather than numbers. | Batch import numeric enum data and apply a single style for display. | Generate Excel files from .NET applications without manual mapping.
// AI Prompts: Generate C# code that uses Aspose.Cells ImportArrayList to insert enum integer values and applies a custom number format to display the enum names. | Show how to build a Style with a custom number format for three enum members and apply it to a range after ImportArrayList. | Explain how to extend the custom number format string for enums with more than three members in Aspose.Cells.

using System;
using System.Collections;
using Aspose.Cells;

enum Status
{
    Unknown = 0,
    Active = 1,
    Inactive = 2
}

// Shows how to create an ArrayList of enum integers, import it into a worksheet using ImportArrayList, define a Style with a custom number format that maps 0‑2 to "Unknown", "Active", "Inactive", apply the style to the range, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the cells collection
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Prepare an ArrayList with enum numeric values
        ArrayList enumValues = new ArrayList();
        enumValues.Add((int)Status.Unknown);
        enumValues.Add((int)Status.Active);
        enumValues.Add((int)Status.Inactive);

        // Import the values vertically starting at cell A1 (row 0, column 0)
        cells.ImportArrayList(enumValues, 0, 0, true);

        // Create a style with a custom number format that maps numbers to enum names
        Style enumStyle = workbook.CreateStyle();
        // Positive;Negative;Zero;Text format sections – we use three sections for 0,1,2
        enumStyle.Custom = "\"Unknown\";\"Active\";\"Inactive\"";

        // Apply the style to the imported range (A1:A3)
        for (int i = 0; i < enumValues.Count; i++)
        {
            cells[i, 0].SetStyle(enumStyle);
        }

        // Save the workbook
        workbook.Save("EnumDisplay.xlsx");
    }
}
