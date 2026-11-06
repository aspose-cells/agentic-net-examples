// Title: How to set the default date format to dd-mmm-yyyy in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change workbook default date format to dd-mmm-yyyy | set custom date pattern for new Excel file with Aspose.Cells .NET | default style date format example in C# Aspose.Cells | apply dd-MMM-yyyy date format to all cells in a workbook using Aspose.Cells | modify default style custom property in Aspose.Cells C#
// Tags: Aspose.Cells default style custom date format | C# workbook default date pattern | Aspose.Cells change default date format Excel | apply dd-mmm-yyyy format Aspose.Cells | Excel workbook default date format .NET

using Aspose.Cells;

// The example creates a new Workbook, retrieves its DefaultStyle, sets the Custom property to "dd-mmm-yyyy", reassigns the modified style back to the workbook, and saves the file as WorkbookWithCustomDateFormat.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Retrieve the default style of the workbook
        Style defaultStyle = workbook.DefaultStyle;

        // Set a custom date format: day-month abbreviation-year (e.g., 05-Jul-2023)
        defaultStyle.Custom = "dd-mmm-yyyy";

        // Apply the modified style back to the workbook's default style
        workbook.DefaultStyle = defaultStyle;

        // Save the workbook to a file
        workbook.Save("WorkbookWithCustomDateFormat.xlsx");
    }
}
