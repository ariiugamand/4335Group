using System.IO;
using System.Windows;
using System.Linq;
using OfficeOpenXml;
using System.Globalization;
using System;

namespace _4335Project
{
    public partial class _4335_Spiridonova : Window
    {
        public _4335_Spiridonova()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }


        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            ImportData();
        }
        private void ExportButton_Click(Object sender, RoutedEventArgs e)
        {
            ExportData();
        }

        private void ImportData()
        {
            var filePath = @"C:/Users/Ксения/source/repos/1.xlsx";

            if (!File.Exists(filePath))
            {
                MessageBox.Show("Файл не найден.");
                return;
            }
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    MessageBox.Show("Файл Excel не содержит листов.");
                    return;
                }

                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension?.Rows ?? 0;

                if (rowCount == 0)
                {
                    MessageBox.Show("Лист Excel пуст.");
                    return;
                }

                using (var context = new AppDbContext())
                {
                    for (int row = 2; row <= rowCount; row++)
                    {
                        try
                        {
                            var idText = worksheet.Cells[row, 1].Text.Trim();
                            var costText = worksheet.Cells[row, 3].Text.Trim();

                            if (string.IsNullOrWhiteSpace(idText) || string.IsNullOrWhiteSpace(costText))
                            {
                                Console.WriteLine($"Пустое значение в строке {row}");
                                continue;
                            }

                            int id;
                            if (!int.TryParse(idText, out id))
                            {
                                Console.WriteLine($"Некорректное значение ID в строке {row}: {idText}");
                                continue;
                            }

                            decimal cost;
                            if (!decimal.TryParse(costText.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out cost))
                            {
                                Console.WriteLine($"Некорректное значение стоимости в строке {row}: {costText}");
                                continue;
                            }

                            var service = new Service
                            {
                                Id = id,
                                Name = worksheet.Cells[row, 2].Text,
                                Cost = cost
                            };

                            context.Services.Add(service);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Ошибка в строке {row}: {ex.Message}");
                        }
                    }

                    context.SaveChanges();
                }
            }

            MessageBox.Show("Данные успешно импортированы!");
        }
        private void ExportData()
        {
            var filePath = @"C:/Users/Ксения/source/repos/exported_data.xlsx";
            var file = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(file))
            {
                using (var context = new AppDbContext())
                {
                    var services = context.Services.OrderBy(s => s.Cost).ToList();
                    var groupedServices = services.GroupBy(s => s.Name);

                    foreach (var group in groupedServices)
                    {
                        var worksheet = package.Workbook.Worksheets.Add(group.Key);

                        worksheet.Cells[1, 1].Value = "Id";
                        worksheet.Cells[1, 2].Value = "Название услуги";
                        worksheet.Cells[1, 3].Value = "Стоимость";

                        int row = 2;
                        foreach (var service in group)
                        {
                            worksheet.Cells[row, 1].Value = service.Id;
                            worksheet.Cells[row, 2].Value = service.Name;
                            worksheet.Cells[row, 3].Value = service.Cost;
                            row++;
                        }
                    }
                }
                package.Save();
            }
            MessageBox.Show("Данные успешно экспортированы!");
        }
    }
}
