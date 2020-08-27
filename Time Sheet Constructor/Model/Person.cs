using System;
using System.Collections.Generic;

namespace Time_Sheet_Constructor.Model
{
    /// <summary>
    /// Оператор
    /// </summary>
    public class Person
    {
        /// <summary>
        /// Имя
        /// </summary>
        public string FirstName { get; set; }

        /// <summary>
        /// Отчество
        /// </summary>
        public string MiddleName { get; set; }

        /// <summary>
        /// Фамилия
        /// </summary>
        public string LastName { get; set; }

        /// <summary>
        /// Табельный номер
        /// </summary>
        public int EmployeeId { get; set; }

        /// <summary>
        /// Расписание
        /// </summary>
        public List<Day> Schedule { get; set; }

        /// <summary>
        /// Первый рабочий день
        /// </summary>
        public int FirstWorkDay => GetFirstWorkDay();        

        /// <summary>
        /// Дата приема сотрудника
        /// </summary>
        public DateTime DateOfReceipt { get; set; }

        public Person()
        {
            Schedule = new List<Day>();            
        }

        /// <summary>
        /// Получаем номер первого рабочего дня
        /// </summary>
        /// <returns></returns>
        private int GetFirstWorkDay()
        {
            var number = 0;
            foreach (var day in Schedule)
            {
                if (day.ScheduledDay)
                {
                    number = day.Number;
                    break;
                }
            }

            return number;
        }

        /// <summary>
        /// Фамилия + Имя
        /// </summary>
        /// <returns></returns>
        public string GetShortName()
        {
            return $"{LastName} {FirstName}";
        }

        /// <summary>
        /// Фамилия + Имя + Отчество
        /// </summary>
        /// <returns></returns>
        public string GetFullName()
        {
            return $"{LastName} {FirstName} {MiddleName}";
        }

    }
}
