using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace WinFormsApp1
{
    public partial class Processor
    {
        [GeneratedRegex(@"([а-яА-ЯёЁ]{2,4})/(\d{2})")]
        private static partial Regex TagRegex();
        
        [GeneratedRegex(@"\bкл_([0-9]{1,2})\b")]
        private static partial Regex ClusterRegex();

        [GeneratedRegex(@"\bрег_([0-9]{1,2})\b")]
        private static partial Regex RICRegex();

        [GeneratedRegex(@"\b([а-яА-ЯёЁ]{2,4}_[а-яА-ЯёЁ]{2,4})_(\d{2})")]
        private static partial Regex RICGroupRegex();

        private readonly bool _multifile;
        private readonly bool _RIC;
        private readonly bool _certification;

        public Processor(bool flag1, bool flag2, bool flag3)
        { 
            _multifile = flag1;
            _RIC = flag2;
            _certification = flag3;
        }
        
        public List<Student> ReadFromFile(string filePath)
        {
            var regFN = new Regex(@"ФИО");
            var indexFN = 0;

            var regTag = new Regex(@"Категории");
            var indexTag = 0;

            var regRegion = new Regex(@"Регион");
            var indexRegion = 0;

            var regCohort = new Regex(@"Поток");
            var indexCohort = 0;

            var regWS = new Regex(@".*\(Значение\)$");
            var indexWS = new Dictionary<int, string>();

            var result = new List<Student>();

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            var headerRow = worksheet.Row(1); // Colomns title row

            foreach (var cell in headerRow.CellsUsed())
            {
                if (regFN.IsMatch(cell.GetString()))
                {
                    indexFN = cell.Address.ColumnNumber;
                }

                if (regTag.IsMatch(cell.GetString()))
                {
                    indexTag = cell.Address.ColumnNumber;
                }

                if (regRegion.IsMatch(cell.GetString()))
                {
                    indexRegion = cell.Address.ColumnNumber;
                }

                if (regCohort.IsMatch(cell.GetString()))
                {
                    indexCohort = cell.Address.ColumnNumber;
                }

                if (regWS.IsMatch(cell.GetString()))
                {
                    indexWS.Add(cell.Address.ColumnNumber, cell.GetValue<string>());
                }
            }

            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

            foreach (var row in rows)
            {
                var fullName = row.Cell(indexFN).GetString();


                string tag = string.Empty;
                int group = 0;
                string? cluster = string.Empty;

                var tagRaw = row.Cell(indexTag).GetString()
                    .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

                var groupRegex = _RIC ? RICGroupRegex() : TagRegex();
                var clusterRegex = _RIC ? RICRegex() : ClusterRegex();

                var groupMatch = tagRaw
                    .Select(m => groupRegex.Match(m))
                    .FirstOrDefault(m => m.Success);

                if (groupMatch is { Success : true })
                {
                    tag = groupMatch.Groups[1].Value;
                    group = int.Parse(groupMatch.Groups[2].Value);
                }

                cluster = tagRaw
                    .Select(c => clusterRegex.Match(c))
                    .FirstOrDefault(m => m.Success)?.Value;

                var region = row.Cell(indexRegion).GetString();
                
                var cohort = row.Cell(indexCohort).GetString();
                if (cohort == "0.00")
                {
                    continue;
                }

                var results = new Dictionary<string, string>();

                foreach (var cell in row.CellsUsed())
                {
                    foreach (var item in indexWS.Keys)
                    {
                        if (cell.Address.ColumnNumber == item)
                        {
                            results.Add(indexWS[item], cell.GetString());
                        }
                    }
                }

                result.Add(new Student
                {
                    FullName = fullName,
                    Cluster = cluster,
                    Group = group,
                    Tag = tag,
                    Region = region,
                    Cohort = cohort,
                    Results = ConvertToResults(results)
                });
            }

            return result;
        }

        public void WriteToFile(string filePath, List<Student> Students)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Filtered");

            // Write header
            int col = 1;
            worksheet.Cell(1, col++).Value = "ФИО";
            worksheet.Cell(1, col++).Value = _RIC ? "Код региона" : "Кластер";
            worksheet.Cell(1, col++).Value = "Группа";
            worksheet.Cell(1, col++).Value = "Поток";
            worksheet.Cell(1, col++).Value = "Регион";
            worksheet.Cell(1, col++).Value = "Кол-во\nсданных\nработ";

            (int, int) WorkIndexes;
            WorkIndexes.Item1 = col;

            var filterKey = "Аттестация";

            foreach(var workName in _certification 
                ? Students[1].FilterKeys(filterKey).Keys 
                : Students[1].Results.Keys)
            {
                worksheet.Cell(1, col++).Value = workName;
            }
            WorkIndexes.Item2 = col - 1;

            int row = 2;
            foreach(var stud in Students)
            {
                col = 1;
                worksheet.Cell(row, col++).Value = stud.FullName;
                worksheet.Cell(row, col++).Value = stud.Cluster;
                worksheet.Cell(row, col++).Value = _RIC ? stud.GetFullGroup('_') : stud.GetFullGroup('/');
                worksheet.Cell(row, col++).Value = stud.Cohort;
                worksheet.Cell(row, col++).Value = stud.Region;
                worksheet.Cell(row, col++).Value = _certification ? stud.FilterKeys(filterKey).Ratio : stud.GetRatio();

                foreach (var workValue in _certification
                    ? stud.FilterKeys(filterKey).Values
                    : stud.Results.Values)
                {
                    var cell = worksheet.Cell(row, col++);
                    cell.Value = workValue.ToString().Replace("_", " ");

                    cell.Style.Fill.BackgroundColor = workValue switch
                    {
                        WorkStatus.Зачтено => XLColor.FromHtml("#C6EFCE"),
                        WorkStatus.На_проверке => XLColor.FromHtml("#FFEB9D"),
                        _ => XLColor.FromHtml("#FFC7CE")
                    };
                }
                row++;
            }

            var table = worksheet.Range(1, 1, row - 1, WorkIndexes.Item2).AsTable();
            table.Name = "Учащиеся";
            table.ShowAutoFilter = true;
            table.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            table.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);

            worksheet.Columns(1, 6).AdjustToContents();

            worksheet.Row(1).AdjustToContents();
            worksheet.Row(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Row(1).Style.Font.Bold = true;

            worksheet.Columns(1, WorkIndexes.Item1 - 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Columns(WorkIndexes.Item1, WorkIndexes.Item2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Columns(WorkIndexes.Item1, WorkIndexes.Item2).Width = 20;

            workbook.SaveAs(filePath);
        }

        private static WorkStatus ConvertToWorkStatus(string status)
        {
            return status switch
            {
                "Зачтено" => WorkStatus.Зачтено,
                "Не зачтено" or "Не_зачтено" => WorkStatus.Не_зачтено,
                "Не сдано" or "Не_сдано" => WorkStatus.Не_сдано,
                "На проверке" or "На_проверке" => WorkStatus.На_проверке,
                "Не доступно" or "Не_доступно" or "Недоступно" => WorkStatus.Недоступно,
                _ => throw new ArgumentException($"Неизвестный статус: {status}")
            };
        }

        private static Results ConvertToResults(Dictionary<string, string> stringDict)
        {
            return new Results(
                stringDict.ToDictionary(
                    kvp => kvp.Key,
                    kvp => ConvertToWorkStatus(kvp.Value)
                    ));
        }

        public void WriteFilesToDir(string filePath, List<Student> Students)
        {
            // Делаем суперсловарь: (
            // Ключ - Кластер,
            // Значение - Словарь:
            // (Ключ - Тег, Значение - Список учащихся)
            // )
            var filteredStudents = Students
                .GroupBy(student => student.Cluster)
                .ToDictionary(
                    clusterGroup => clusterGroup.Key,
                    clusterGroup => clusterGroup
                        .GroupBy(student => student.Tag)
                        .ToDictionary(
                            tagGroup => tagGroup.Key,
                            tagGroup => tagGroup.ToList()
                        )
                );

            // Бежим по верхнему словарю с ключ-кластером
            foreach(var (clusterKey, dict) in filteredStudents)
            {
                var cluster = clusterKey;

                // Создаем папку кластера
                string clusterDirectoryPath = System.IO.Path.Combine(filePath, cluster);
                Directory.CreateDirectory(clusterDirectoryPath);

                // Бежим по внутреннему словарю с ключ-тегом
                foreach (var (tagKey, list) in dict)
                {
                    var tag = tagKey;

                    // Формируем имя файла
                    string FileName = $"{tag}_{cluster}";
                    // Путь к файлу
                    string tagFilePath = System.IO.Path.Combine(clusterDirectoryPath, FileName);

                    // Сборка таблицы
                    using var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add($"{FileName}");

                    int col = 1;
                    worksheet.Cell(1, col++).Value = "ФИО";
                    worksheet.Cell(1, col++).Value = _RIC ? "Код региона" : "Кластер";
                    worksheet.Cell(1, col++).Value = "Группа";
                    worksheet.Cell(1, col++).Value = "Регион";
                    worksheet.Cell(1, col++).Value = "Кол-во\nсданных\nработ";

                    (int, int) WorkIndexes;
                    WorkIndexes.Item1 = col;
                    foreach(var workName in list[0].Results.Keys)
                    {
                        worksheet.Cell(1, col++).Value = workName;
                    }
                    WorkIndexes.Item2 = col - 1;

                    int row = 2;
                    foreach (var stud in list)
                    {
                        col = 1;
                        worksheet.Cell(row, col++).Value = stud.FullName;
                        worksheet.Cell(row, col++).Value = stud.Cluster;
                        worksheet.Cell(row, col++).Value = _RIC ? stud.GetFullGroup('_') : stud.GetFullGroup('/');
                        worksheet.Cell(row, col++).Value = stud.Region;
                        worksheet.Cell(row, col++).Value = stud.GetRatio();

                        foreach (var workValue in stud.Results.Values)
                        {
                            var cell = worksheet.Cell(row, col++);
                            cell.Value = workValue.ToString().Replace("_", " ");

                            cell.Style.Fill.BackgroundColor = workValue switch
                            {
                                WorkStatus.Зачтено => XLColor.FromHtml("#C6EFCE"),
                                WorkStatus.На_проверке => XLColor.FromHtml("#FFEB9D"),
                                _ => XLColor.FromHtml("#FFC7CE")
                            };
                        }
                        row++;
                    }

                    // row - 1 костыль, т.к. после последней строки у нас идет еще 1 инкремент
                    var table = worksheet.Range(1, 1, row - 1, WorkIndexes.Item2).AsTable();

                    table.Name = "Учащиеся";
                    table.ShowAutoFilter = true;
                    table.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    table.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);

                    worksheet.Columns(1, 5).AdjustToContents();

                    worksheet.Row(1).AdjustToContents();
                    worksheet.Row(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Row(1).Style.Font.Bold = true;

                    worksheet.Columns(1, WorkIndexes.Item1 - 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    worksheet.Columns(WorkIndexes.Item1, WorkIndexes.Item2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Columns(WorkIndexes.Item1, WorkIndexes.Item2).Width = 20;

                    workbook.SaveAs(tagFilePath + ".xlsx");
                }
            }
        }
    }
}
