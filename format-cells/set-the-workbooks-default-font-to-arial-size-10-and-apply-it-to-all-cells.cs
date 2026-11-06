// Title: How to set the default workbook font to Arial 10pt and apply it to every cell with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set workbook default font to Arial 10 point | apply a global font style to all cells in an Excel file using Aspose.Cells | C# code to change default cell font and propagate to existing worksheets in Aspose.Cells | using StyleFlag to apply default style to every cell in Aspose.Cells workbook | set default style font for new and existing worksheets Aspose.Cells .NET
// Tags: default workbook font Aspose.Cells .NET | global cell style application Aspose.Cells | StyleFlag All property usage | set Arial 10pt font workbook default | propagate default style across worksheets

using Aspose.Cells;

// The example creates a Workbook, sets its DefaultStyle font to Arial 10 pt, configures a StyleFlag with All=true, and applies the default style to every cell in all worksheets before saving the file.
class Program
{
    static void Main()
    {
        // create workbook (use provided create rule)
        Workbook workbook = new Workbook();

        // set default font to Arial, size 10
        workbook.DefaultStyle.Font.Name = "Arial";
        workbook.DefaultStyle.Font.Size = 10;

        // apply the default style to all cells in every worksheet
        StyleFlag flag = new StyleFlag();
        flag.All = true;

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Cells.ApplyStyle(workbook.DefaultStyle, flag);
        }

        // save workbook (use provided save rule)
        workbook.Save("output.xlsx");
    }
}
