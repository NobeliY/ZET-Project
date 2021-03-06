using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using ZET_Project.Classes.CSV;
using ZET_Project.Classes.Manager;

namespace ZET_Project.Classes.Employees
{
    public class NewEmployee
    {
        protected internal string Name;
        protected internal string Surname;
        protected internal string Login;
        protected internal string Password;
        protected internal string Post;
        

        public NewEmployee SetNewEmployee(string name, string surname, string post, string login,
            string password )
        {
            return new NewEmployee()
                {Name = name, Surname = surname, Post = post, Login = login, Password = password};
        }
    }
    public class AuthLog
    {
        protected internal string? Login { get; set; }
        protected internal string? Password { get; set; }

        public AuthLog(string? login, string? password)
        {
            this.Login = login;
            this.Password = password;
        }
    }

    public class LowPerson
    {
        public LowPerson LowPersons(string post, int hours)
        {
            Post = post;
            Hours = hours;
            switch (Post)
            {
                case "Freelancer":
                    power = 1;
                    sum = 1000;
                    summed = sum * 160;
                    break;
                case "Accountant":
                    power = 2;
                    sum = 750;
                    summed = sum * 160;
                    break;
                case "Director":
                    power = 1;
                    sum = 1250;
                    summed = sum * 160;
                    break;
            }

            return new LowPerson(){Hours = Hours, Post = Post, power = power,sum = sum,summed = summed};
        }

        protected internal int power; //Уровень дополнительной оплаты
        protected internal int sum; // Сама зарплата (по часам)
        protected internal int summed; //Зар.Плата за месяц
        protected internal string Post;
        protected internal int Hours;
        
    }
    public class Person
    {
        public Person(string name, string surname, string post)
        {
            Name = name;
            Surname = surname;
            Post = post;
        }
        protected internal string Name;
        protected internal string Surname;
        protected internal string Post;
    }
    


    public class Freelancer
    {
        private string _message = $"Выберите желаемое действие:\n" +
                                           //$" (1). Добавить сотрудника.\n" +
                                           //$" (2). Просмотреть отчет по всем сотрудникам.\n" +
                                           $" (1). Просмотреть отчет.\n" + $" (2). Добавить дополнительные часы.\n" + $" (0). Выход из программы";

        public virtual void SendInformation()
        {
            rst:
            Console.Clear();
            Console.WriteLine(_message);
            ConsoleKeyInfo key = Console.ReadKey();
            switch (key.KeyChar)
            {
                case '1':
                    GetInformation(CsvRead.Post);
                    break;
                case '2':
                    Console.Clear();
                    Console.WriteLine("Введите дату:");
                    var dateTime = Console.ReadLine();
                    Console.WriteLine("Введите количество добавляемых часов: ");
                    var hours = int.Parse(Console.ReadLine() ?? string.Empty);
                    AddTime(dateTime, hours);
                    break;
                case '0':
                    Console.Clear();
                    break;
                default:
                    Console.WriteLine("Этой кнопки нет в данном списке! \n Попробуйте нажать другую.");
                    goto rst;

            }
        }

        protected void GetInformation(string? tableList)
        {
            var process = new Process();
            process.StartInfo = new ProcessStartInfo($"Employee{tableList}.xlsx")
            {
                UseShellExecute = true
            };
            process.Start();
            process.WaitForExit();
        }

        protected virtual void AddTime(string dateTime, int hours)
        {
            var todayDates = DateTime.Today.ToString("d", new CultureInfo("fr-FR"));
            var todayDateSplit = todayDates.Split('/');
            var todayDate = int.Parse(todayDateSplit[0]);
            var dateSplit = dateTime.Split('/');
            var dateSplitDay = int.Parse(dateSplit[0]);
            if (todayDate >= dateSplitDay)
            {
                if (todayDate - dateSplitDay > 2)
                {
                    Console.WriteLine("Напишите, за что добавляется время.");
                    var note = Console.ReadLine();
                    ExcelManager excelManager = new();
                    excelManager.AddHours(initials: EmployeeManager.Initials,
                        date: dateTime, hours: hours, tableList: CsvRead.Post, note: note);
                }
                else
                {
                   Console.WriteLine($"Вы не можете добавить время за {dateTime}! Прошло больше 2 дней!");
                }
            }
        }
    }

    public class Accountant : Freelancer
    {
        protected override void AddTime(string dateTime, int hours)
        {
            Console.WriteLine("Напишите, за что добавляется время.");
            var note = Console.ReadLine();
            ExcelManager excelManager = new();
            excelManager.AddHours(EmployeeManager.Initials,
                dateTime, hours, tableList: CsvRead.Post, note: note);
        }
    }
    

    public class Director : Freelancer
    {
        private string _message = $"Выберите желаемое действие:\n" +
                                  $" (1). Просмотреть отчет.\n" + $" (2). Добавить дополнительные часы.\n" +
                                  $"(3). Посмотреть подробный отчет по сотруднику за период. \n" +
                                  $"(4). Добавить сотрудника. \n" +
                                  $" (0). Выход из программы";
        static Dictionary<string, ExcelPerson> persons;
        static Dictionary<string, LowPerson> personalsReportData = new();
        private protected void GetReportEmployee(string name, string dateArray)
        {
            int sum = 0;

            
            ExcelManager excelManager = new();
            if (name.Contains("All") || name.Contains("Все"))
            {
                excelManager.GetReportEmployeeArray("All","All");
                //ExcelManager.ExcelPersonAlls
                string fullyName;
                foreach (var (key, value) in CsvRead.GetPersonals())
                {
                    fullyName = value.Name + " " + value.Surname;
                    personalsReportData.Add(fullyName,
                        new LowPerson().LowPersons(CsvRead.GetPost(fullyName),0));
                    Console.WriteLine(personalsReportData[fullyName]);
                }
                string report = $"Отчет по сотрудникам: \n";
                int hoursSum = 0;
                long sums = 0;
                foreach (var variablePersonalsReportData in personalsReportData)
                {
                    foreach (var variableExcelPersonAll in ExcelManager.ExcelPersonAlls)
                    {
                        if (variablePersonalsReportData.Key.Equals(variableExcelPersonAll.Value.employeeName))
                        {
                            personalsReportData[variablePersonalsReportData.Key].Hours += variableExcelPersonAll.Value.hours;
                        }
                        
                    }
                    hoursSum += variablePersonalsReportData.Value.Hours;

                    sums += variablePersonalsReportData.Value.power * variablePersonalsReportData.Value.Hours *
                        variablePersonalsReportData.Value.sum + variablePersonalsReportData.Value.summed;

                    report +=
                        $"{variablePersonalsReportData.Key} отработал {variablePersonalsReportData.Value.Hours.ToString()}. Сумма к выдаче: {variablePersonalsReportData.Value.summed.ToString()} + {(variablePersonalsReportData.Value.sum * variablePersonalsReportData.Value.power * variablePersonalsReportData.Value.Hours).ToString()} = {(variablePersonalsReportData.Value.power * variablePersonalsReportData.Value.Hours * variablePersonalsReportData.Value.sum + variablePersonalsReportData.Value.summed).ToString()} рублей. \n";
                }

                report += $"Итого отработано {hoursSum.ToString()}, сумма к выплате {sums}";
                Console.Clear();
                Console.WriteLine(report);

            }
            else
            {
                switch (CsvRead.GetPost(name))
                {
                    case "Freelancer":
                        sum = 1000;
                        break;
                    case "Accountant":
                        sum = 750;
                        break;
                    case "Director":
                        sum = 1250;
                        break;
                }
                string[] dateArrays = dateArray.Split('-');
                var tableList = CsvRead.GetPost(name); 
                excelManager.GetReportEmployeeArray(name, tableList);
                bool detect = false;
                foreach (var (key, value) in ExcelManager.ExcelPersons)
                {
                    if (value.date.Equals(dateArrays[0]))
                    {
                        detect = true;
                    }

                    if (detect)
                    {
                        if (value.date.Equals(dateArrays[1]))
                        {
                            break;
                        }
                        else
                        {
                            persons.Add(key,value);
                        }
                    }
                }

                string report = $"Отчет по сотруднику: [{name}] за период с [{dateArrays[0]}] по [{dateArrays[1]}] \n ";
                int hoursSum = 0;
                if (CsvRead.GetPost(name) == "Director")
                {
                    int summed = 20000;
                    foreach (var (key,value) in persons)
                    {
                        hoursSum += value.hours;
                    }

                    report +=
                        $"Есть переработка, составляющая {hoursSum}. Дополнительная оплата в размере {summed} зачтена!";
                    Console.Clear();
                    Console.WriteLine(report);
                }
                else
                {
                    foreach (var person in persons)
                    {
                        hoursSum += person.Value.hours;
                        report += $"{person.Key}, {person.Value.hours.ToString()} часов, {person.Value.note} \n";
                    }

                    int summed = sum * hoursSum;
                    report += $"Итого: {hoursSum} часов, заработано: {summed.ToString()} рублей.";
                    Console.Clear();
                    Console.WriteLine(report);
                }
                
            }
            
        }
        protected override void AddTime(string? dateTime, int hours)
        {
            Console.Clear();
            RestartNameRead:
            Console.WriteLine("Введите Фамилию и Имя сотрудника (Важно: вводить раздельно фамилию и имя) :");
            string employeeNames = Console.ReadLine();
            if (employeeNames != null)
            {
                if (CsvRead.GetPost(employeeNames) == "Not State")
                {
                    Console.WriteLine("Такого сотрудника нет в предприятии");
                    goto RestartNameRead;
                }
                Console.WriteLine("Введите примечание (за что добавляете время?): ");
                var note = Console.ReadLine();
                ExcelManager excelManager = new();
                excelManager.AddHours(employeeNames, dateTime, hours, tableList: CsvRead.GetPost(employeeNames), note: note);

            }
            else
            {
                Console.WriteLine("Вы ввели пустую строку! Пожалуйста повторите попытку.");
                goto RestartNameRead;
            }
        }

        private bool DetectEmployee(string[] fullyName)
        {
            foreach (var variablePerson in CsvRead.GetPersonals())
            {
                
                if ((variablePerson.Value.Name.Equals(fullyName[0]) && variablePerson.Value.Surname.Equals(fullyName[1]))
                    || (variablePerson.Value.Name.Equals(fullyName[1]) && variablePerson.Value.Surname.Equals(fullyName[0])))
                    return true;
            }
            return false;
        }
        public void SetNewEmployees()
        {
            
            Dictionary<int, NewEmployee> newEmployees = new();
            
            int nums = 0;
            RSTNumsSet:
            Console.Clear();
            Console.WriteLine("Введите количество сотрудников: ");
            if (int.TryParse(Console.ReadLine(), out nums))
            {
                int count = CsvRead.GetLastId();
                for (int i = 0; i < nums; i++)
                {
                    RestartReadFullyName:
                    Console.Clear();
                    Console.WriteLine("Введите Имя и Фамилию: ");
                    string[] fullyName = Console.ReadLine().Split(' ');
                    if (DetectEmployee(fullyName))
                    {
                        Console.WriteLine(
                            "Сотрудник с таким Именем и Фамилией уже существует в списке. Повторите попытку ввода еще раз!");
                        Console.ReadLine();
                        goto RestartReadFullyName;
                    }

                    Console.WriteLine("Введите должность (Доступно: Director, Accountant, Freelancer): ");
                    string post = Console.ReadLine();
                    Console.WriteLine("Установите логин для входа: ");
                    string login = Console.ReadLine();
                    Console.WriteLine("Установите пароль для входа: ");
                    string password = Console.ReadLine();
                    newEmployees.Add(count, new NewEmployee().SetNewEmployee(fullyName[0],
                        fullyName[1], post,
                        login, password));
                    count++;

                }
                CsvRead.AddEmployeeListFile(newEmployees);
            }
            else
            {
                Console.WriteLine("Нужно ввести число, а не текст!");
                Thread.Sleep(500);
                goto RSTNumsSet;
            }



        }

        public override void SendInformation()
        {
            rst:
            Console.Clear();
            Console.WriteLine(_message);
            ConsoleKeyInfo key = Console.ReadKey();
            switch (key.KeyChar)
            {
                case '1':
                    GetInformation(tableList: CsvRead.Post);
                    break;
                case '2':
                    Console.Clear();
                    Console.WriteLine("Введите дату:");
                    var dateTime = Console.ReadLine();
                    Console.WriteLine("Введите количество добавляемых часов: ");
                    var hours = int.Parse(Console.ReadLine() ?? string.Empty);
                    AddTime(dateTime, hours);
                    break;
                case '3':
                    Console.Clear();
                    RestartEmployeeSet:
                    Console.WriteLine("Введите Фамилию и Имя Сотрудника (Если хотите получить отчет по всем сотрудника, Введите <Все> или <All>) : ");
                    string employeeName = Console.ReadLine();
                    if (employeeName != null)
                    {
                        if (employeeName.Contains("All") || employeeName.Contains("Все"))
                        {
                            GetReportEmployee("All", "All");
                        }
                        else
                        {
                            RestartDateArraySet:
                            Console.WriteLine("Введите период (Например. 01/01/2021-08/01/2021): ");
                            string dateArray = Console.ReadLine();
                            if (dateArray != null)
                            {
                                GetReportEmployee(name: employeeName, dateArray: dateArray);
                            }
                            else
                            {
                                Console.WriteLine("Вы ввели пустую строку! Повторите попытку ввода ещё раз.");
                                goto RestartDateArraySet;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Вы ввели пустую строку! Повторите попытку ввода ещё раз.");
                        goto RestartEmployeeSet;
                    }
                    
                    break;
                case '4':
                    Console.Clear();
                    SetNewEmployees();
                    break;
                case '0':
                    Console.WriteLine();
                    break;
                default:
                    Console.WriteLine("Этой кнопки нет в данном списке! \n Попробуйте нажать другую.");
                    goto rst;
            }
        }
    }
}