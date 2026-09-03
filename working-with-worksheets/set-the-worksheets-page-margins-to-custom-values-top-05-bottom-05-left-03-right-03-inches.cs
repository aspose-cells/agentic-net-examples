// Title: Set custom worksheet page margins (0.5" top/bottom, 0.3" left/right) with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells to set a worksheet's top and bottom margins to 0.5 inches and left/right margins to 0.3 inches, then save the workbook. | Demonstrate how to modify the PageSetup.Margin properties in Aspose.Cells to apply custom inch‑based margins to an Excel sheet.
// Common Searches: Aspose.Cells C# example for setting worksheet margins in inches | how to change top bottom left right page margins with Aspose.Cells .NET | custom page margin values for Excel file using Aspose.Cells PageSetup | C# code to set worksheet margins to 0.5 and 0.3 inches with Aspose.Cells
// Tags: Aspose.Cells worksheet page margins | PageSetup margin properties C# | custom Excel margins using Aspose.Cells | set worksheet margins inches .NET

using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, sets the PageSetup margins to 0.5 inches for top and bottom and 0.3 inches for left and right, and saves the file as CustomMargins.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set custom page margins (values are in inches)
        sheet.PageSetup.TopMargin = 0.5;
        sheet.PageSetup.BottomMargin = 0.5;
        sheet.PageSetup.LeftMargin = 0.3;
        sheet.PageSetup.RightMargin = 0.3;

        // Save the workbook to a file
        workbook.Save("CustomMargins.xlsx");
    }
}
