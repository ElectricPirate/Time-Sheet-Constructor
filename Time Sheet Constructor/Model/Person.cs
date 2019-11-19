namespace Time_Sheet_Constructor.Model
{
    public class Person
    {
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public int EmployeeId { get; set; }
        public Day[] Schedule { get; set; }

        public Person() { }

        public Person(string firstName, string lastName, Day[] schedule)
        {
            FirstName = firstName;
            LastName = lastName;
            Schedule = schedule;
        }

        public override string ToString()
        {
            return $"{LastName} {MiddleName} {FirstName}";
        }
    }
}
