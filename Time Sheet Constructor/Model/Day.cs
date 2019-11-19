namespace Time_Sheet_Constructor.Model
{/// <summary>
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
        public string Absense { get; set; }

        public Day() { }

        public Day(double allWorkTime, double nightWorkTime, bool dayOff, bool sickDay, bool vacationDay,
            bool unpaidLeave, bool educationalLeave, string absense)
        {
            AllWorkTime = allWorkTime;
            NightWorkTime = nightWorkTime;
            DayOff = dayOff;
            SickDay = sickDay;
            VacationDay = vacationDay;
            UnpaidLeave = unpaidLeave;
            EducationalLeave = educationalLeave;
            Absense = absense;
        }
    }
}
