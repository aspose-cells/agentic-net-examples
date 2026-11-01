// Title: How to set a workbook’s default date format to "dd-mmm-yyyy" with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set default date format for entire workbook | apply custom date style globally in Excel using Aspose.Cells | C# Aspose.Cells default style with dd-mmm-yyyy date format | change workbook default date pattern to dd-mmm-yyyy in Aspose.Cells | set default cell style for dates in Aspose.Cells C#
// Tags: default date format Aspose.Cells C# | custom date style workbook Aspose.Cells | Workbook.DefaultStyle property Aspose.Cells | Style.Custom date pattern Aspose.Cells | Excel date formatting Aspose.Cells C#

using Aspose.Cells;
using System;

// The example creates a Workbook, defines a Style with the custom date format "dd-mmm-yyyy", assigns this Style to the workbook's DefaultStyle, and saves the file, ensuring all date cells inherit the specified format.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Create a style with the desired custom date format
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "dd-mmm-yyyy";

        // Set this style as the workbook's default style
        workbook.DefaultStyle = dateStyle;

        // Save the workbook to a file
        workbook.Save("WorkbookWithCustomDateFormat.xlsx");
    }
}
