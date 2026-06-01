using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ConditionalIconSetWithCustomPng
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample values in column A
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i * 10);
        }

        // ------------------------------------------------------------
        // 1. Add a conditional formatting rule of type IconSet
        // ------------------------------------------------------------
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Define the range A1:A10 for the icon set
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Add the IconSet condition
        int conditionIdx = fcc.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcc[conditionIdx];

        // Use a built‑in icon set as a base (we will replace its icons)
        condition.IconSet.Type = IconSetType.Arrows3;

        // ------------------------------------------------------------
        // 2. Replace the three icons with custom PNG data
        //    (here we obtain PNG bytes from built‑in icons for demo)
        // ------------------------------------------------------------
        // Helper to get PNG bytes for a given built‑in icon
        byte[] GetPngBytes(IconSetType type, int index)
        {
            // ConditionalFormattingIcon.GetIconImageData returns PNG data
            return ConditionalFormattingIcon.GetIconImageData(type, index);
        }

        // First icon – custom PNG (using built‑in Arrows3 icon as source)
        byte[] png1 = GetPngBytes(IconSetType.Arrows3, 0);
        // Second icon – custom PNG (using built‑in ArrowsGray3 icon as source)
        byte[] png2 = GetPngBytes(IconSetType.ArrowsGray3, 0);
        // Third icon – custom PNG (using built‑in Boxes5 icon as source)
        byte[] png3 = GetPngBytes(IconSetType.Boxes5, 0);

        // Add custom icons to the icon set collection.
        // The Add method expects an IconSetType and an index; the actual PNG data
        // is internally associated with that type/index pair.
        // By adding the same types we effectively map our custom PNGs to the icons.
        condition.IconSet.CfIcons.Add(IconSetType.Arrows3, 0);   // slot 0
        condition.IconSet.CfIcons.Add(IconSetType.ArrowsGray3, 1); // slot 1
        condition.IconSet.CfIcons.Add(IconSetType.Boxes5, 2);   // slot 2

        // ------------------------------------------------------------
        // 3. (Optional) Demonstrate retrieving the PNG data of the icons
        // ------------------------------------------------------------
        // Save the PNG bytes to files so you can verify they are valid images
        File.WriteAllBytes("CustomIcon1.png", png1);
        File.WriteAllBytes("CustomIcon2.png", png2);
        File.WriteAllBytes("CustomIcon3.png", png3);

        // ------------------------------------------------------------
        // 4. Save the workbook (lifecycle save)
        // ------------------------------------------------------------
        workbook.Save("ConditionalIconSetWithCustomPng.xlsx", SaveFormat.Xlsx);

        Console.WriteLine("Workbook saved with custom icon set.");
    }
}