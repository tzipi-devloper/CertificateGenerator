using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificateGenerator
{
    internal class Employee
    {
        public string FirstName { get; }
        public string LastName { get; }
        public string Department { get; }
        public double FinalScore { get; }

        public string FullName => $"{FirstName} {LastName}";

        public Employee(string[] columns)
        {
            FirstName = Capitalize(columns[0].Trim());
            LastName = Capitalize(columns[1].Trim());
            Department = columns[2].Trim();

            double.TryParse(columns[3], out double theory);
            double.TryParse(columns[4], out double practical);
            FinalScore = (practical * 0.6) + (theory * 0.4);
        }

        private static string Capitalize(string input) =>
            string.IsNullOrEmpty(input) ? input : char.ToUpper(input[0]) + input.Substring(1).ToLower();
    }
}

