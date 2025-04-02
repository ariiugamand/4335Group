using System.IO;
using System.Windows;
using System.Linq;
using OfficeOpenXml;
using System.Globalization;
using System;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO.Packaging;
using Aspose.Words.Rendering;
using Xceed.Words.NET;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data.SqlClient;
using System.Data;
namespace _4335Project
{
    public partial class _4335_Spiridonova : System.Windows.Window
    {
        private List<Service> services = new List<Service>();
        public _4335_Spiridonova()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void ImportJson_Click(object sender, RoutedEventArgs e)
        {
            ImportJson();
        }
        private void ExportWord_Click(object sender, RoutedEventArgs e)
        {
            ExportJson();
        }
       
        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            ImportData();
        }
        private void ExportButton_Click(Object sender, RoutedEventArgs e)
        {
            ExportData();
        }
        private void ImportJson()
        {
            try
            {
                string json = File.ReadAllText(@"C:/Users/Ксения/source/repos/1.json");
                services = JsonConvert.DeserializeObject<List<Service>>(json);
                using (var connection = new SqlConnection(@"Server=DESKTOP-EJ9IKF0\MSSQLSERVER01; Database=TestDB; Integrated Security=True;"))
                {
                    connection.Open();
                    foreach (var service in services)
                    {
                        var command = new SqlCommand(
                            "INSERT INTO Services (IdServices, NameServices,TypeOfService, Cost) VALUES (@id, @name, @type, @cost)",
                            connection
                        );
                        command.Parameters.AddWithValue("@id", service.IdServices);
                        command.Parameters.AddWithValue("@name", service.NameServices);
                        command.Parameters.AddWithValue("@type", service.TypeOfService);
                        command.Parameters.Add("@cost", SqlDbType.Decimal).Value = service.Cost;
                        command.ExecuteNonQuery();
                    }
                }
                using (var context = new AppDbContext())
                {
                    context.SaveChanges();
                }
                MessageBox.Show("Данные успешно импортированы");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
        private void ExportJson()
        {
            try
            {
                if (services == null || !services.Any())
                {
                    MessageBox.Show("Сначала загрузите данные через импорт!");
                    return;
                }
                var groupedServices = services
                    .GroupBy(s => s.TypeOfService)
                    .OrderBy(g => g.Key)
                    .ToDictionary(g => g.Key, g => g.OrderBy(s => s.Cost).ToList());

                string filePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "ServicesReport.docx"
                );
                if (File.Exists(filePath)) File.Delete(filePath);
                using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    foreach (var group in groupedServices)
                    {
                        body.Append(new Paragraph(
                    new Run(
                        new Text($"Категория: {group.Key}"),
                        new RunProperties(new Bold())
                    )
                ));
                        Table table = new Table(
                            new TableProperties(
                                new TableBorders(
                                    new TopBorder() { Val = BorderValues.Single, Size = 4 },
                                    new BottomBorder() { Val = BorderValues.Single, Size = 4 },
                                    new LeftBorder() { Val = BorderValues.Single, Size = 4 },
                                    new RightBorder() { Val = BorderValues.Single, Size = 4 }
                                )
                            )
                        );
                        table.Append(new TableRow(
                            CreateCell("ID", true),
                            CreateCell("Название услуги", true),
                            CreateCell("Стоимость", true)
                        ));

                        foreach (var service in group.Value)
                        {
                            table.Append(new TableRow(
                         CreateCell(service.IdServices.ToString()),
                         CreateCell(service.NameServices),
                         CreateCell(service.Cost.ToString("C"))
                     ));
                        }

                        body.Append(table);
                        body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    }

                    MessageBox.Show($"Документ успешно сохранен: {filePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private TableCell CreateCell(string text, bool isHeader = false)
        {
            TableCell cell = new TableCell();
            Paragraph paragraph = new Paragraph();

            if (isHeader)
            {
                paragraph.Append(new Run(
                    new Text(text),
                    new RunProperties(new Bold())
                ));
            }
            else
            {
                paragraph.Append(new Run(new Text(text)));
            }

            cell.Append(paragraph);
            return cell;
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
                                IdServices = id,
                                NameServices = worksheet.Cells[row, 2].Text,
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
                    var groupedServices = services.GroupBy(s => s.NameServices);

                    foreach (var group in groupedServices)
                    {
                        var worksheet = package.Workbook.Worksheets.Add(group.Key);

                        worksheet.Cells[1, 1].Value = "Id";
                        worksheet.Cells[1, 2].Value = "Название услуги";
                        worksheet.Cells[1, 3].Value = "Стоимость";

                        int row = 2;
                        foreach (var service in group)
                        {
                            worksheet.Cells[row, 1].Value = service.IdServices;
                            worksheet.Cells[row, 2].Value = service.NameServices;
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
