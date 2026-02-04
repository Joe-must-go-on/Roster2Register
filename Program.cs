using ClosedXML.Excel;
using ExcelLogic;

// gets the number of the current month as a string
// uses it to make sure i only use the dates in the current month to get shift times
DateTime currentDate = DateTime.Now;
int monthNum = currentDate.Month - 1;//will be running after the month has ended, so we need to get last month
string correctMonthNum = monthNum.ToString();
if (correctMonthNum.Length == 1)
{
    correctMonthNum = "0" + correctMonthNum;
}
string currentYear = currentDate.Year.ToString();

// file paths for the schedule and attendance register excel files
string ScheduleFilePath = @$"C:\Users\User\Desktop\Wimpy Register\Lone {currentYear}-{correctMonthNum}.xlsx";
string RegisterFilePath = @"C:\Users\User\Desktop\Wimpy Register\Attendance Register.xlsx";

//cells where the dates are located in the schedule file
string scheduleDatesRow = "5"; 
string[] scheduleDatesCol ={
    "D",
    "H",
    "L",
    "P",
    "T",
    "X",
    "AB"
};

string[] scheduleWorkSheets =
{
    "Week1",
    "Week2",
    "Week3",
    "Week4",
    "Week5",
    "Week6"
};



using (ExcelHandler registerWorkBook = new ExcelHandler(RegisterFilePath, 1)){
    using (ExcelHandler ScheduleWorkBook = new ExcelHandler(ScheduleFilePath))
    {
        foreach (string week in scheduleWorkSheets){
            foreach (string date in scheduleDatesCol)
            {
                string dayOfMonth;
                var cellDate = ScheduleWorkBook.GetCellData(date+scheduleDatesRow, week) as string;
                if (cellDate[5..7] == correctMonthNum)
                {
                    string day = cellDate[8..10];
                    if (day[0] == '0')
                    {
                        dayOfMonth = day.Replace("0", string.Empty);
                    }
                    else
                    {
                        dayOfMonth = day;
                    }
                    //only gets the list for one day
                    var (employees, clockIns, clockOuts) = ScheduleWorkBook.GetShiftTimes(
                    columnLetter: date, 
                    startRow: 6, 
                    endRow: 52, 
                    workSheet: week);

                    int index = employees.IndexOf("Mandoza");
                    if (index != -1)//index -1 means it was not found
                    {
                        employees[index] = "Reolebogile";
                    }
                    registerWorkBook.WriteToRoster(employees, clockIns, clockOuts, dayOfMonth);
                }
                else
                {

                }
            }
        }
    }
    registerWorkBook.Save();
}

//go through the roster file and make sure all the lunch times fit within the shift times
//how to iterate only through clock in and clock out columns?
//if clock in < lunch start < lunch end < clock out then its valid
//if not then adjust lunch times to fit within shift times
//if clock out - clock in < 5 hours then no lunch break

