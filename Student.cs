using System.ComponentModel.DataAnnotations;

namespace WinFormsApp1
{
    public class Student
    {
        [Required]
        public string FullName { get; set; } = string.Empty;
        public string Cluster { get; set; } = string.Empty;
        public int Group { get; set; }
        public string Region { get; set; } = string.Empty;
        public string Cohort { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
        public Dictionary<string, WorkStatus> Results { get; set; }
        
        public int GetRatio()
        {
            if (Results == null)
            {
                return 0;
            }
            return Results.Count(pair => pair.Value == WorkStatus.Зачтено);
        }

        public string? GetFullGroup()
        {
            if (Tag == null || Group == 0)
            {
                return null;
            }
            return $"{Tag}/{Group:D2}";
        }

    }

    public enum WorkStatus
    {
        Зачтено,
        Не_зачтено,
        Не_сдано,
        На_проверке,
        Не_доступно,
        Недоступно
    }
}
