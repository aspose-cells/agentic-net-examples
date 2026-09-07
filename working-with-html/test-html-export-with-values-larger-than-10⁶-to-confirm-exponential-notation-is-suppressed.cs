// Title: Suppress scientific notation when exporting numbers over one million to HTML with Aspose.Cells for .NET
// AI Prompts: Generate C# code that writes values greater than 1,000,000 into an Aspose.Cells workbook, applies a custom numeric format to keep the numbers as plain integers, and saves the sheet as an HTML file. | Show how to configure Aspose.Cells Style.Custom with the format string "0" to avoid exponential notation in the HTML export.
// Common Searches: Aspose.Cells how to keep large numbers from showing scientific notation in HTML output | C# export Excel to HTML with plain integer formatting using Aspose.Cells | prevent exponential notation for values over 1 million when saving workbook as HTML Aspose.Cells | custom number format 0 for HTML export in Aspose.Cells .NET example
// Tags: Aspose.Cells HTML custom numeric style | disable exponential display in HTML export | format cells as integer for HTML output | C# Aspose.Cells set style custom 0 | HTML export large integer values Aspose.Cells

using Aspose.Cells;

// The example creates a workbook, inserts two numbers larger than one million, applies a custom "0" number format to each cell to keep them as plain integers, and saves the workbook as an HTML file, ensuring scientific notation is not used.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert values larger than 10⁶
        sheet.Cells["A1"].PutValue(1234567);
        sheet.Cells["A2"].PutValue(9876543);

        // Apply a plain integer number format to suppress scientific notation
        Style plainStyle = workbook.CreateStyle();
        plainStyle.Custom = "0";
        sheet.Cells["A1"].SetStyle(plainStyle);
        sheet.Cells["A2"].SetStyle(plainStyle);

        // Export the workbook to HTML (save rule)
        workbook.Save("LargeNumbers.html", SaveFormat.Html);
    }
}
