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


//hour list = [11,12,13,14,15]
//minute list = [00,15,30,45]
// create a timeonly object from the hour and minute lists, then instert into lunch start, then +30minutes for lunch end



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
                    registerWorkBook.NameSearch(employees, clockIns, clockOuts, dayOfMonth);
                }
                else
                {

                }
            }
        }
    }
    registerWorkBook.Save();
}


