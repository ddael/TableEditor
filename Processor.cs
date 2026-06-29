using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace WinFormsApp1
{
    public partial class Processor
    {
        [GeneratedRegex(@"([а-яА-ЯёЁ]{2,4})/(\d{2})")]
        private static partial Regex TagRegex();
        
        [GeneratedRegex(@"\bкл_([1-9][0-9]{0,2})\b")]
        private static partial Regex ClusterReg();
        
        public static List<Student> ReadFromFile(string filePath)
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

                var tagRaw = row.Cell(indexTag).GetString()
                    .Split(", ")
                    .Select(s => s.Trim())
                    .ToList();
                var match = tagRaw
                    .Select(m => TagRegex().Match(m))
                    .FirstOrDefault(m => m.Success);
                string tag = string.Empty;
                int group = 0;
                if (match is { Success : true })
                {
                    tag = match.Groups[1].Value;
                    group = int.Parse(match.Groups[2].Value);
                }
                var cluster = tagRaw
                    .Select(c => ClusterReg().Match(c))
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
                    Results = ConvertDictionary(results)
                });
            }

            return result;
        }

        public static void WriteToFile(string filePath, List<Student> Students)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Filtered");

            // Write header
            worksheet.Cell(1, 1).Value = "ФИО";
            worksheet.Cell(1, 2).Value = "Кластер";
            worksheet.Cell(1, 3).Value = "Группа";
            worksheet.Cell(1, 4).Value = "Поток";
            worksheet.Cell(1, 5).Value = "Регион";
            worksheet.Cell(1, 6).Value = "Кол-во\nсданных\nработ";
            List<string> Headers = Students[1].Results.Keys.ToList();

            for (int i = 0; i < Headers.Count; i++)
            {
                worksheet.Cell(1, 6 + (i + 1)).Value = Headers[i];
            }

            // Write data
            for (int i = 0; i < Students.Count; i++)
            {
                worksheet.Cell(i + 2, 1).Value = Students[i].FullName;
                worksheet.Cell(i + 2, 2).Value = Students[i].Cluster;
                worksheet.Cell(i + 2, 3).Value = Students[i].GetFullGroup();
                Console.WriteLine(Students[i].GetFullGroup());
                worksheet.Cell(i + 2, 4).Value = Students[i].Cohort;
                worksheet.Cell(i + 2, 5).Value = Students[i].Region;
                worksheet.Cell(i + 2, 6).Value = Students[i].GetRatio();
                List<WorkStatus> StudStat = Students[i].Results.Values.ToList();
                for (int j = 0; j < StudStat.Count; j++)
                {
                    var cell = worksheet.Cell(i + 2, 6 + (j + 1));
                    cell.Value = StudStat[j].ToString().Replace("_", " ");

                    if (StudStat[j] == WorkStatus.Зачтено)
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#C6EFCE");
                    }
                    else if (StudStat[j] == WorkStatus.На_проверке)
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFEB9D");
                    }
                    else
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC7CE");
                    }
                }
            }

            var table = worksheet.Range(1, 1, Students.Count + 1, Headers.Count + 6).AsTable();
            table.Name = "Учащиеся";
            table.ShowAutoFilter = true;
            table.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            table.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);

            worksheet.Columns(1, 6).AdjustToContents();

            worksheet.Row(1).AdjustToContents();
            worksheet.Row(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Row(1).Style.Font.Bold = true; 

            worksheet.Columns(1, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            int lastColumn = 7 + (Students.Count > 0 ? Students[0].Results.Count : 0);
            worksheet.Columns(7, lastColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Columns(7, lastColumn).Width = 20;

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

        private static Dictionary<string, WorkStatus> ConvertDictionary(Dictionary<string, string> stringDict)
        {
            return stringDict.ToDictionary(
                kvp => kvp.Key,
                kvp => ConvertToWorkStatus(kvp.Value)
                );
        }

        public static void WriteFilesToDir(string filePath, List<Student> Students)
        {
            // Формируем словарь: Ключ - Кластер, Занчение - Список учащихся в кластере
            var groupedByCluster = Students
                .GroupBy(student => student.Cluster)
                .ToDictionary(grouped => grouped.Key, grouped => grouped.ToList());

            foreach(var clusterKvp in groupedByCluster)
            {
                var cluster = clusterKvp.Key;
                var studentsInCluster = clusterKvp.Value;

                // Создаем папку кластера
                string clusterDirectoryPath = System.IO.Path.Combine(filePath, cluster);
                Directory.CreateDirectory(clusterDirectoryPath);

                // Формируем IEnumerable список по группам внутри кластера
                var groupsInCluster = studentsInCluster
                    .GroupBy(student => student.Group);

                foreach (var group in groupsInCluster)
                {
                    var groupName = group.Key;
                    var studentsByGroup = group.ToList();

                    // Формируем имя файла
                    string FileName = $"{groupName}_{cluster}";
                    // Путь к файлу
                    string groupFilePath = System.IO.Path.Combine(clusterDirectoryPath, FileName);

                    // Сборка таблицы
                    using var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add($"{cluster}");

                    worksheet.Cell(1, 1).Value = "ФИО";
                    worksheet.Cell(1, 2).Value = "Кластер";
                    worksheet.Cell(1, 3).Value = "Группа";
                    worksheet.Cell(1, 4).Value = "Регион";
                    worksheet.Cell(1, 5).Value = "Кол-во\nсданных\nработ";
                    List<string> Headers = studentsByGroup[1].Results.Keys.ToList();

                    for (int i = 0; i < studentsByGroup.Count; i++)
                    {
                        worksheet.Cell(1, 5 + (i + 1)).Value = Headers[i];
                    }

                    for (int i = 0; i < studentsByGroup.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = Students[i].FullName;
                        worksheet.Cell(i + 2, 2).Value = Students[i].Cluster;
                        worksheet.Cell(i + 2, 3).Value = Students[i].Group;
                        worksheet.Cell(i + 2, 4).Value = Students[i].Region;
                        worksheet.Cell(i + 2, 5).Value = Students[i].GetRatio();
                        List<WorkStatus> StudStat = studentsByGroup[i].Results.Values.ToList();
                        for (int j = 0; j < StudStat.Count; j++)
                        {
                            var cell = worksheet.Cell(i + 2, 5 + (j + 1));
                            cell.Value = StudStat[j].ToString().Replace("_", " ");

                            if (StudStat[j] == WorkStatus.Зачтено)
                            {
                                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#C6EFCE");
                            }
                            else if (StudStat[j] == WorkStatus.На_проверке)
                            {
                                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFEB9D");
                            }
                            else
                            {
                                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC7CE");
                            }
                        }
                    }

                    var table = worksheet.Range(1, 1, studentsByGroup.Count + 1, Headers.Count + 5).AsTable();

                    table.Name = "Учащиеся";
                    table.ShowAutoFilter = true;
                    table.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    table.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);

                    worksheet.Columns(1, 5).AdjustToContents();

                    worksheet.Row(1).AdjustToContents();
                    worksheet.Row(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Row(1).Style.Font.Bold = true;

                    worksheet.Columns(1, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    int lastColumn = 6 + (Students.Count > 0 ? Students[0].Results.Count : 0);
                    worksheet.Columns(6, lastColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Columns(6, lastColumn).Width = 20;

                    workbook.SaveAs(groupFilePath);
                }
            }
        }
    }
}
