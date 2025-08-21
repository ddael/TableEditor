using System.ComponentModel.DataAnnotations;

namespace WinFormsApp1
{
    public class Student
    {
        [Required]
        public string FullName { get; set; } = string.Empty;
        public string Group { get; set; } = string.Empty;
        public string Region { get; set; } = string.Empty;
        public string Cohort { get; set; } = string.Empty;
        public Dictionary<string, WorkStatus> Results { get; set; } 
        public WorkStatus Status { get; set; }
        
        public int GetRatio()
        {
            if (Results == null)
            {
                return 0;
            }
            return Results.Count(pair => pair.Value == WorkStatus.Зачтено);
        }

    }

    public enum WorkStatus
    {
        Зачтено,
        Не_зачтено,
        Не_сдано,
        На_проверке,
        Не_доступно
    }
}
