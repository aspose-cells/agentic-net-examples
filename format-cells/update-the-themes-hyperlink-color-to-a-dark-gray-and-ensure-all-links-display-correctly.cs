using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHyperlinkThemeDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a sample hyperlink to demonstrate the theme effect
            // Parameters: cell name, row offset, column offset, hyperlink address
            sheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

            // Update the theme's hyperlink color to a dark gray
            // ThemeColorType.Hyperlink corresponds to the hyperlink color in the theme
            // Color.DarkGray provides a standard dark gray color
            workbook.SetThemeColor(ThemeColorType.Hyperlink, Color.DarkGray);

            // Save the workbook; all hyperlinks will now use the dark gray theme color
            workbook.Save("HyperlinkThemeDemo.xlsx");
        }
    }
}