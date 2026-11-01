// Title: Apply a solid light-gray background and bold font to the first worksheet row (header) with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Style with a solid light-gray fill and bold text, then applies it to the first row of a worksheet using Aspose.Cells. | Write a C# snippet that configures a StyleFlag to apply all style attributes and uses Row.ApplyStyle to format the header row in an Excel workbook.
// Common Searches: Aspose.Cells C# set header row background color to light gray | How to make the first row bold and gray in an Excel file with Aspose.Cells | C# apply solid fill pattern to an entire row using Aspose.Cells | Using StyleFlag to apply multiple style attributes to a row in Aspose.Cells
// Tags: apply solid fill to header row Aspose.Cells | header row bold font Aspose.Cells | StyleFlag apply all attributes C# | set background color light gray Aspose.Cells | format first worksheet row Excel .NET

using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, defines a style with a solid light-gray background and bold font, applies the style to the first worksheet row (header) using a StyleFlag, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Create a style for the header row
        Style headerStyle = workbook.CreateStyle();
        // Set background color to light gray
        headerStyle.ForegroundColor = Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;
        // Optional: make the header text bold
        headerStyle.Font.IsBold = true;

        // Define which style attributes to apply
        StyleFlag flag = new StyleFlag();
        flag.All = true; // apply all style attributes

        // Apply the style to the first row (header row)
        Row headerRow = sheet.Cells.Rows[0];
        headerRow.ApplyStyle(headerStyle, flag);

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
