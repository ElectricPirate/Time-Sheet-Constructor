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

        public Person()
        {
            Schedule = new List<Day>();
        }

        public override string ToString()
        {
            return $"{LastName} {FirstName}";
        }

    }
}
