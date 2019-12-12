using System.ComponentModel;

namespace Time_Sheet_Constructor.Model
{
    /// <summary>
    /// День
    /// </summary>
    public class Day
    {
        /// <summary>
        /// Номер дня месяца
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// Все рабочее время
        /// </summary>
        public double AllWorkTime { get; set; }

        /// <summary>
        /// Всерхурочное рабочее время
        /// </summary>
        public double OverTime { get; set; }

        /// <summary>
        /// Ночное рабочее время
        /// </summary>
        public double NightWorkTime { get; set; }

        /// <summary>
        /// Выходной
        /// </summary>
        public bool DayOff { get; set; }

        /// <summary>
        /// Больничный
        /// </summary>
        public bool SickDay { get; set; }

        /// <summary>
        /// Отпуск ежегодный
        /// </summary>
        public bool VacationDay { get; set; }

        /// <summary>
        /// Отпуск дополнительный
        /// </summary>
        public bool UnpaidLeave { get; set; }

        /// <summary>
        /// Отпуск учебный
        /// </summary>
        public bool EducationalLeave { get; set; }

        /// <summary>
        /// Неявка
        /// </summary>
        public bool Truancy { get; set; }

        /// <summary>
        /// Отпуск по беременности и родам
        /// </summary>
        public bool MaternityLeave { get; set; }

        public bool IsCrossing => GetCrossings();

        private bool GetCrossings()
        {
            return AllWorkTime != 0 && SickDay || VacationDay || UnpaidLeave || EducationalLeave || Truancy;
        }

    }
}
