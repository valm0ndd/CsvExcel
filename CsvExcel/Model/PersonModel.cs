using CsvHelper.Configuration.Attributes;

namespace CsvExcel.Model
{
    public class PersonModel
    {
        [Name("Index")]
        public int Index { get; set; }

        [Name("User Id")]
        public string UserId { get; set; }

        [Name("First Name")]
        public string FirstName { get; set; }

        [Name("Last Name")]
        public string LastName { get; set; }

        [Name("Sex")]
        public string Sex { get; set; }

        [Name("Email")]
        public string Email { get; set; }

        [Name("Phone")]
        public string Phone { get; set; }

        [Name("Date of birth")]
        public DateTime DateOfBirth { get; set; }

        [Name("Job Title")]
        public string JobTitle { get; set; }
    }
}