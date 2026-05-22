using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultiThreadedChartExport
{
    public class ChartExportDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook with memory‑efficient settings.
                using (Workbook workbook = new Workbook())
                {
                    workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

                    // Access the first worksheet and enable multi‑thread reading.
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;
                    cells.MultiThreadReading = true;

                    // Populate sample data.
                    int rows = 100;
                    cells[0, 0].PutValue("Category");
                    for (int i = 1; i <= rows; i++)
                    {
                        cells[i, 0].PutValue($"Item {i}");
                        cells[i, 1].PutValue(i * 10);   // English
                        cells[i, 2].PutValue(i * 12.5); // French
                        cells[i, 3].PutValue(i * 9.8);  // German
                    }

                    // Prepare chart parameters.
                    int chartCount = 3;
                    int[] dataColumns = { 1, 2, 3 };
                    string[] locales = { "en", "fr", "de" };
                    string outputFolder = Path.Combine(Environment.CurrentDirectory, "ChartExports");
                    Directory.CreateDirectory(outputFolder);

                    // Create charts.
                    Chart[] charts = new Chart[chartCount];
                    for (int i = 0; i < chartCount; i++)
                    {
                        int chartIndex = sheet.Charts.Add(ChartType.Column, 5 + i * 15, 0, 20 + i * 15, 8);
                        Chart chart = sheet.Charts[chartIndex];

                        string valueRange = dataColumns[i] switch
                        {
                            2 => $"C2:C{rows + 1}",
                            3 => $"D2:D{rows + 1}",
                            _ => $"B2:B{rows + 1}"
                        };

                        chart.NSeries.Add(valueRange, true);
                        chart.NSeries.CategoryData = $"A2:A{rows + 1}";
                        chart.Title.Text = $"Sales ({locales[i].ToUpperInvariant()})";

                        chart.Calculate(); // optional layout calculation
                        charts[i] = chart;
                    }

                    // Export charts to PDF in parallel.
                    Parallel.ForEach(
                        Enumerable.Range(0, chartCount),
                        new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                        index =>
                        {
                            try
                            {
                                Chart chart = charts[index];
                                string locale = locales[index];
                                string pdfPath = Path.Combine(outputFolder, $"Chart_{locale}.pdf");
                                chart.ToPdf(pdfPath);
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine($"Error exporting chart {locales[index]}: {ex.Message}");
                            }
                        });

                    // Save workbook for verification.
                    string workbookPath = Path.Combine(outputFolder, "WorkbookWithCharts.xlsx");
                    workbook.Save(workbookPath, SaveFormat.Xlsx);
                }

                Console.WriteLine("Chart export completed. Files are located at: " + Path.Combine(Environment.CurrentDirectory, "ChartExports"));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Unexpected error: " + ex.Message);
            }
        }
    }

    // Entry point required by the runtime.
    public static class Program
    {
        public static void Main(string[] args)
        {
            ChartExportDemo.Run();
        }
    }
}