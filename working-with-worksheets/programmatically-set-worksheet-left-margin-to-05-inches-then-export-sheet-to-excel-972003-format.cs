// Title: Set worksheet left margin to 0.5 inches and save as Excel 97‑2003 (XLS) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set left page margin to 0.5 inches | How to export a workbook to XLS format with Aspose.Cells .NET | Convert inches to points for page margins in Aspose.Cells | Save workbook as Excel 97‑2003 file using Aspose.Cells C# | Programmatically configure page setup margins with Aspose.Cells
// Tags: set worksheet left margin Aspose.Cells | save workbook as XLS using Aspose.Cells | page setup margin points to inches conversion | SaveFormat.Excel97To2003 example C# | Aspose.Cells page setup API C#

using Aspose.Cells;

// The example creates a new Workbook, accesses its first Worksheet, sets the left page margin to 0.5 inches by converting inches to points (0.5 × 72), and then saves the workbook as an Excel 97‑2003 XLS file using the SaveFormat.Excel97To2003 option.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the left margin to 0.5 inches (1 inch = 72 points)
        sheet.PageSetup.LeftMargin = 0.5 * 72;

        // Export the workbook to Excel 97‑2003 format (XLS)
        workbook.Save("output.xls", SaveFormat.Excel97To2003);
    }
}
