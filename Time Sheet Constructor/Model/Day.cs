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

        private double _allWorkTime;

        /// <summary>
        /// Все рабочее время
        /// </summary>
        public double AllWorkTime 
        { 
            get
            {
                return _allWorkTime;
            }
            set
            {
                if (Math.Truncate(value * 10) == 0)
                    _allWorkTime = 0;
                else
                    _allWorkTime = Math.Round(value, 1);
            }
        }

        private double _overTime;
        /// <summary>
        /// Сверхурочное рабочее время
        /// </summary>
        public double OverTime
        {
            get
            {
                return _overTime;
            }
            set
            {
                if (Math.Truncate(value * 10) == 0)
                    _overTime = 0;
                else
                    _overTime = Math.Round(value, 1);
            }
        }

        private double _nightWorkTime;

        /// <summary>
        /// Ночное рабочее время
        /// </summary>
        public double NightWorkTime
        {
            get
            {
                return _nightWorkTime;
            }
            set
            {
                if (Math.Truncate(value * 10) == 0)
                    _nightWorkTime = 0;
                else
                    _nightWorkTime = Math.Round(value, 1);
            }
        }

        /// <summary>
        /// Выходной
        /// </summary>
        public bool DayOff { get; set; }

        private string _sickDay;
        /// <summary>
        /// Больничный
        /// </summary>
        public string SickDay 
        {
            get
            {
                return _sickDay;
            }
            set
            {
                _sickDay = $"Б/{value}";
            }
        }

        private string _vacationDay;
        /// <summary>
        /// Отпуск ежегодный
        /// </summary>
        public string VacationDay
        {
            get
            {
                return _vacationDay;
            }
            set
            {
                _vacationDay = $"ОТ/{value}";
            }
        }

        private string _unpaidLeave;
        /// <summary>
        /// Отпуск дополнительный
        /// </summary>
        public string UnpaidLeave
        {
            get
            {
                return _unpaidLeave;
            }
            set
            {
                _unpaidLeave = $"ДО/{value}";
            }
        }

        private string _educationLeave;
        /// <summary>
        /// Отпуск учебный
        /// </summary>
        public string EducationalLeave
        {
            get
            {
                return _educationLeave;
            }
            set
            {
                _educationLeave = $"У/{value}";
            }
        }

        private string _truancy;
        /// <summary>
        /// Неявка
        /// </summary>
        public string Truancy
        {
            get
            {
                return _truancy;
            }
            set
            {
                _truancy = $"НН/{value}";
            }
        }

        private string _hooky;
        /// <summary>
        /// Прогул
        /// </summary>
        public string Hooky
        {
            get
            {
                return _hooky;
            }
            set
            {
                _hooky = $"ПР/{value}";
            }
        }

        private string _maternityLeave;
        /// <summary>
        /// Отпуск по беременности и родам
        /// </summary>
        public string MaternityLeave
        {
            get
            {
                return _maternityLeave;
            }
            set
            {
                _maternityLeave = $"ОЖ/{value}";
            }
        }

        private string _paidDayOff;
        /// <summary>
        /// Оплачиваемый выходной. Привет Рома!
        /// </summary>
        public string PaidDayOff
        {
            get
            {
                return _paidDayOff;
            }
            set
            {
                _paidDayOff = $"ОВ/{value}";
            }
        }

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
