using System;
using System.Collections.Generic;
using System.Linq;

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

        public static Person ParseName(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                throw new ArgumentException("ФИО должно содержать минимум фамилию и имя.", nameof(fullName));
            }

            var names = fullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (names.Length < 2)
            {
                throw new ArgumentException("ФИО должно содержать минимум фамилию и имя.", nameof(fullName));
            }

            return new Person
            {
                LastName = names[0],
                FirstName = names[1],
                MiddleName = string.Join(" ", names.Skip(2))
            };
        }

        public bool IsSamePerson(string fullName)
        {
            try
            {
                var person = ParseName(fullName);
                return LastName.Equals(person.LastName) && FirstName.Equals(person.FirstName);
            }
            catch (ArgumentException)
            {
                return false;
            }
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
            return string.IsNullOrWhiteSpace(MiddleName) ? GetShortName() : $"{LastName} {FirstName} {MiddleName}";
        }

    }
}
