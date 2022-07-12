using System;
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

        private double allWorkTime;

        /// <summary>
        /// Все рабочее время
        /// </summary>
        public double AllWorkTime 
        { 
            get
            {
                return allWorkTime;
            }
            set
            {
                if (Math.Truncate(value * 10) == 0)
                    allWorkTime = 0;
                else
                    allWorkTime = Math.Round(value, 1);
            }
        }

        private double overTime;
        /// <summary>
        /// Сверхурочное рабочее время
        /// </summary>
        public double OverTime
        {
            get
            {
                return overTime;
            }
            set
            {
                if (Math.Truncate(value * 10) == 0)
                    overTime = 0;
                else
                    overTime = Math.Round(value, 1);
            }
        }

        private double nightWorkTime;

        /// <summary>
        /// Ночное рабочее время
        /// </summary>
        public double NightWorkTime
        {
            get
            {
                return nightWorkTime;
            }
            set
            {
                if (Math.Truncate(value * 10) == 0)
                    nightWorkTime = 0;
                else
                    nightWorkTime = Math.Round(value, 1);
            }
        }

        /// <summary>
        /// Выходной
        /// </summary>
        public bool DayOff { get; set; }

        /// <summary>
        /// Больничный
        /// </summary>
        public string SickDay { get; set; }

        /// <summary>
        /// Отпуск ежегодный
        /// </summary>
        public string VacationDay { get; set; }

        /// <summary>
        /// Отпуск дополнительный
        /// </summary>
        public string UnpaidLeave { get; set; }

        /// <summary>
        /// Отпуск учебный
        /// </summary>
        public string EducationalLeave { get; set; }

        /// <summary>
        /// Неявка
        /// </summary>
        public string Truancy { get; set; }

        /// <summary>
        /// Прогул
        /// </summary>
        public string Hooky { get; set; }

        /// <summary>
        /// Отпуск по беременности и родам
        /// </summary>
        public string MaternityLeave { get; set; }

        /// <summary>
        /// Оплачиваемый выходной. Привет Рома!
        /// </summary>
        public string PaidDayOff { get; set; }

        /// <summary>
        /// Запланирован ли день
        /// </summary>
        public bool ScheduledDay => IsScheduledDay();

        private bool IsScheduledDay() => AllWorkTime != default || OverTime != default || NightWorkTime != default || DayOff || SickDay != default || VacationDay != default || UnpaidLeave != default || EducationalLeave != default || Truancy != default || MaternityLeave != default || PaidDayOff != default || Hooky != default;

        /// <summary>
        /// Пересечение рабочих часов и отсутствия
        /// </summary>
        public bool Crossing => GetCrossings();       

        private bool GetCrossings()
        {
            //return AllWorkTime != default && (SickDay != default || VacationDay != default || UnpaidLeave != default || EducationalLeave != default || Truancy != default || MaternityLeave != default || PaidDayOff != default || Hooky != default);

            var count = 0;

            if (AllWorkTime != default)
                count++;            
            if (OverTime != default)
                count++;
            if (SickDay != default)
                count++;
            if (VacationDay != default)
                count++;
            if (UnpaidLeave != default)
                count++;
            if (EducationalLeave != default)
                count++;
            if(Truancy != default)
                count++;
            if(MaternityLeave != default)
                count++;
            if(PaidDayOff != default)
                count++;
            if(Hooky != default)
                count++;

            return count>1?true:false;
        }

    }
}
