// Title: How to set the default date format to dd-mmm-yyyy in an Aspose.Cells workbook using C#
// AI Prompts: Apply a custom number format 'dd-mmm-yyyy' to the workbook's default style and save the workbook. | Configure Aspose.Cells to display dates in the 'dd-mmm-yyyy' pattern for all new cells by modifying the default style.
// Common Searches: Aspose.Cells C# set workbook default date format to dd-mmm-yyyy | change default date pattern for new Excel files using Aspose.Cells | how to customize default style date format in Aspose.Cells .NET | set global date format in Aspose.Cells workbook programmatically
// Tags: set workbook date pattern Aspose.Cells | modify workbook date pattern C# | global Excel date pattern Aspose.Cells | default workbook style customization .NET | date display pattern Aspose.Cells

using Aspose.Cells;

// Creates a new workbook, changes its default style to use the custom date pattern "dd-mmm-yyyy", and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the create rule)
        Workbook workbook = new Workbook();

        // Change the workbook's default date format to "dd-mmm-yyyy"
        // This modifies the default style's custom number format
        workbook.DefaultStyle.Custom = "dd-mmm-yyyy";

        // Save the workbook (using the save rule)
        workbook.Save("output.xlsx");
    }
}
