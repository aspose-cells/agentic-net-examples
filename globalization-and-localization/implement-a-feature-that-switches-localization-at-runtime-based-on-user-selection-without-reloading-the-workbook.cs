// Title: Change Excel workbook locale at runtime with Aspose.Cells for .NET without reopening the file
// Common Searches: how to switch Excel workbook culture on the fly using Aspose.Cells C# | Aspose.Cells set workbook locale without reloading the workbook | recalculate Excel formulas after changing CultureInfo with Aspose.Cells | runtime localization of an existing .xlsx file in a .NET console app | handle invalid locale errors when updating workbook settings in Aspose.Cells
// Tags: set workbook CultureInfo Aspose.Cells | runtime locale switch Excel .xlsx Aspose.Cells | recalculate formulas after culture change Aspose.Cells | apply localization to existing workbook without reload | Aspose.Cells workbook settings culture update

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to load an existing Excel workbook, assign a new CultureInfo to the workbook's Settings, recalculate all formulas to reflect locale‑dependent changes, and save the file back to its original location—all without reopening the workbook. It includes error handling for missing files and invalid locale strings.
public class LocalizationSwitcher
{
    // Switches the workbook's localization at runtime based on the provided locale string (e.g., "en-US", "fr-FR").
    // The workbook is loaded, culture is changed, formulas are recalculated, and the same file is saved.
    public void SwitchLocalization(string workbookPath, string locale)
    {
        // Verify that the workbook file exists
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

        // Load the existing workbook
        Workbook workbook = new Workbook(workbookPath);

        // Apply the new culture info
        workbook.Settings.CultureInfo = new CultureInfo(locale);

        // Recalculate formulas so that any locale‑dependent functions reflect the new culture
        workbook.CalculateFormula();

        // Save the workbook back to the same file
        workbook.Save(workbookPath);
    }

    // Entry point for the console application
    public static void Main(string[] args)
    {
        try
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: LocalizationSwitcher <workbookPath> <locale>");
                Console.WriteLine("Example: LocalizationSwitcher \"C:\\Files\\Sample.xlsx\" en-US");
                return;
            }

            string workbookPath = args[0];
            string locale = args[1];

            var switcher = new LocalizationSwitcher();
            switcher.SwitchLocalization(workbookPath, locale);

            Console.WriteLine($"Workbook '{workbookPath}' successfully localized to '{locale}'.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.Error.WriteLine($"File error: {fnfEx.Message}");
        }
        catch (CultureNotFoundException cnfEx)
        {
            Console.Error.WriteLine($"Invalid locale: {cnfEx.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
