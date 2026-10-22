// Title: Enable Page Break Preview view for a worksheet in a new workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, sets the first worksheet's ViewType to PageBreakPreview with Aspose.Cells, and saves it as an .xlsx file. | Show how to switch an existing worksheet to Page Break Preview mode using the Aspose.Cells API in a .NET application. | Provide a snippet that demonstrates enabling page break visualization for a worksheet and exporting the result with Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# set worksheet view to page break preview | How to enable page break preview mode in an Excel file using Aspose.Cells for .NET | C# Aspose.Cells change worksheet view type to PageBreakPreview before saving | Page break preview viewtype example Aspose.Cells .NET | Visualize Excel page breaks with Aspose.Cells C# code
// Tags: Aspose.Cells worksheet viewtype pagebreakpreview | C# Aspose.Cells enable page break preview | export workbook with page break preview Aspose | Excel page break visualization Aspose.Cells .NET | set worksheet view to pagebreakpreview C#

using Aspose.Cells;

// Creates a workbook, accesses the first worksheet, sets its ViewType to PageBreakPreview to show page break preview, and saves the file as PageBreakPreview.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable Page Break Preview mode for the worksheet
        worksheet.ViewType = ViewType.PageBreakPreview;

        // Save the workbook to a file
        workbook.Save("PageBreakPreview.xlsx");
    }
}
