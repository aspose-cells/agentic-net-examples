using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkDemo
{
    class Program
    {
        static void Main()
        {
            string inputPath = "InputWithExternalLinks.xlsx";
            string outputPath = "OutputWithUpdatedExternalLinks.xlsx";

            // Ensure the input workbook exists; create a simple one if it does not.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Cells["A1"].Formula = "=SUM(1,2,3)";
                wb.Save(inputPath);
            }

            // Load the workbook that may contain external links.
            Workbook workbook = new Workbook(inputPath);

            // Update the data source of each external link if needed.
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];
                string oldSource = link.DataSource;
                string newSource = oldSource.Replace(
                    "https://oldsharepoint.com/Docs/",
                    "https://newsharepoint.com/SharedDocuments/");

                // Apply the updated data source.
                link.DataSource = newSource;
            }

            // Prepare calculation options with linked data sources.
            CalculationOptions calcOptions = new CalculationOptions();

            int linkCount = workbook.Worksheets.ExternalLinks.Count;
            Workbook[] linkedWorkbooks = new Workbook[linkCount];

            for (int i = 0; i < linkCount; i++)
            {
                string externalPath = workbook.Worksheets.ExternalLinks[i].DataSource;

                // If the external workbook does not exist, create a placeholder.
                if (!File.Exists(externalPath))
                {
                    var placeholder = new Workbook();
                    placeholder.Worksheets[0].Cells["A1"].PutValue(0);
                    placeholder.Save(externalPath);
                }

                linkedWorkbooks[i] = new Workbook(externalPath);
            }

            calcOptions.LinkedDataSources = linkedWorkbooks;

            // Recalculate all formulas using the updated external link information.
            workbook.CalculateFormula(calcOptions);

            // Save the updated workbook.
            workbook.Save(outputPath);
        }
    }
}