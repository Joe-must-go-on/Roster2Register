using DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.CoverPageProps;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelLogic;
// IDisposable makes it automatically close the file
public class ExcelHandler : IDisposable
{
    //Class field that we can use in the methods of this class to manipulate the excel files
    //private means only this class can access it, readonly means it can only be assigned in the constructor
    private readonly XLWorkbook _workbook;
    public ExcelHandler(string filePath)
    {
        // opens the excel file
        _workbook = new XLWorkbook(filePath);
    }

// using a different constructor for the attendance register excel file
// int num is just used to differentiate between the two constructors
    public ExcelHandler(string filePath, int num)
    {
        // opens the excel file
        _workbook = new XLWorkbook(filePath);
        // worksheet used in attendance register excel file
        var attendanceWorkSheet = _workbook.Worksheet("Register");
    }

    /// <summary>
    /// Retrieves the value of a cell from the specified worksheet.
    /// </summary>
    /// <param name="cellNumber">The A1-style cell reference (for example: "A1", "C5").</param>
    /// <param name="workSheet">The name of the worksheet that contains the cell.</param>
    /// <returns>The cell value as a string.</returns>
    /// <remarks>
    /// This method uses ClosedXML to access the workbook. 
    /// The return type is <see cref="object"/>, but the value returned is a string
    /// obtained via <c>GetValue&lt;string&gt;()</c>.
    /// </remarks>
    public object GetCellData(string cellNumber, string workSheet)
    {
        return _workbook.Worksheet(workSheet).Cell(cellNumber).GetValue<string>();
    }

    /// <summary>
    /// Extracts employee names and their clock-in/clock-out times from a specific column range in a worksheet.
    /// </summary>
    /// <param name="columnLetter">The column letter that contains clock-in times (e.g., "C"). The clock-out time is expected in the cell to the right.</param>
    /// <param name="startRow">The first row index to read (inclusive).</param>
    /// <param name="endRow">The last row index boundary (exclusive). Iteration stops at <paramref name="endRow"/> - 1.</param>
    /// <param name="workSheet">The name of the worksheet to read from.</param>
    /// <returns>
    /// A tuple of three lists: <c>Names</c> (employee names from column B), <c>Ins</c> (clock-in times), and <c>Outs</c> (clock-out times),
    /// where time values are returned as strings.
    /// </returns>
    /// <remarks>
    /// A row is included only if both the target cell and its right-adjacent cell contain parsable numeric values (interpreted as doubles).
    /// Employee names are taken from column B of the same row. Non-numeric or empty rows are skipped.
    /// The method uses ClosedXML APIs and assumes the worksheet exists; invalid sheet names may cause exceptions.
    /// </remarks>
    public (List<string> Names, List<string> Ins, List<string> Outs) GetShiftTimes(
        string columnLetter, 
        int startRow, 
        int endRow, 
        string workSheet)
    {
        //will be returning employee, clock in and clock out as lists
        //I will then later iterate through the list to fill in the attendance register
        // creating these to append to the list
        double clockIn;
        double clockOut;
        string employee;

        //creating the lists to return
        List<string> clockInTimes = new List<string>();
        List<string> clockOutTimes = new List<string>();
        List<string> employeeNames = new List<string>();

        //creating the worksheet to iterate through
        // will be iterating through the worksheets
        var sheet = _workbook.Worksheet(workSheet);

        //iterate through all the rows
        for (int row = startRow; row < endRow; row++)
        {
            //gets the current cell location
            string cellAddress = $"{columnLetter}{row}";

            //checks if value of the cell is a double, which means its a clock in time
            //if it  is not a double, then it will definitely not be a clock in or clock out time
            if (double.TryParse(sheet.Cell(cellAddress).Value.ToString(), out clockIn) && 
            double.TryParse(sheet.Cell(cellAddress).CellRight().Value.ToString(), out clockOut))
            //cell right just goes to next cell, only reason for doing both try parses here is to get the conversion out of the way
            // i have to do .value.to string, otherwise ot tries to convert an object which it cant do.
            //so i get the value(an object), then make it a string, then the try parse runs
            {
                //Gets employee name that is on same row as, and always column B, the clock in time
                employee = sheet.Cell($"B{row}").Value.ToString();
                //Console.WriteLine($"Employee: {employee}, Clock In: {clockIn:F2}, CLock Out: {clockOut:F2}, ");
                //adds to the list
                string formattedClockIn = FormatTime(clockIn);
                string formattedClockOut = FormatTime(clockOut);
                employeeNames.Add(employee);
                clockInTimes.Add(formattedClockIn);
                clockOutTimes.Add(formattedClockOut);
            }
            else
            {
                
            }
        }
        return (employeeNames, clockInTimes, clockOutTimes);
    }

    /// <summary>
    /// Formats a numeric time value into a human-readable HH:MM string.
    /// </summary>
    /// <param name="time">
    /// A double representing time, where the integer part is the hour and the fractional part represents minutes in quarter-hour increments
    /// (e.g., 9 becomes "9:00", 9.5 becomes "9:30", 9.25 becomes "9:15", 9.75 becomes "9:45").
    /// </param>
    /// <returns>A string formatted as H:MM or HH:MM.</returns>
    /// <remarks>
    /// The method performs simple string-based normalization assuming the current culture formats decimals with a comma.
    /// It maps fractional values of .25, .5, and .75 to 15, 30, and 45 minutes respectively.
    /// Other fractional values are not handled and may yield unexpected results.
    /// </remarks>
    public string FormatTime(double time)
    {
        //converts the double to string
        string timeString = time.ToString();
        //make sure that the string doesn't have ,75 or something at the end
        if (timeString.Length < 3)
        {// if it is a number like 9 then it makes it 9:00
            timeString = timeString+":00";
        }
        else if (timeString.Length == 3)
        {
            timeString = timeString.Replace(",5",":30");
        }
        //makes sure that the number has something like ,75 at the end
        else if(timeString.Length > 3)
        {
            timeString = timeString.Replace(",",":");
            if (timeString[2..4] == "75")
            {
                timeString = timeString.Replace("75","45");
            }
            else if (timeString[2..4] == "25")
            {
                timeString = timeString.Replace("25","15");
            }
        }
        return timeString;
    }

    /// <summary>
    /// Searches the "Register" worksheet for each employee name in the header row and writes their clock-in and clock-out times for a given day.
    /// </summary>
    /// <param name="employees">List of employee names to search for in row 2.</param>
    /// <param name="clockIns">Parallel list of clock-in times corresponding to <paramref name="employees"/>.</param>
    /// <param name="clockOuts">Parallel list of clock-out times corresponding to <paramref name="employees"/>.</param>
    /// <param name="dayNum">The day number (as a string) used to determine the target row. The method writes to row <c>int.Parse(dayNum) + 4</c>.</param>
    /// <remarks>
    /// The method scans columns 1 through 109 in row 2 to find a cell whose value contains the employee name.
    /// When a match is found, it writes the clock-out time to the matched column and the clock-in time three columns to the left on the computed date row.
    /// After all writes, the workbook is saved.
    /// </remarks>
    public void NameSearch(
        List<string> employees, 
        List<string> clockIns, 
        List<string> clockOuts,
        string dayNum)
    {
        var workSheet = _workbook.Worksheet("Register");// hardcoded because it will always be this worksheet
        int row = 2;//employee names in the attendance register in is row 2
        int startColumn = 1;
        int endColumn = 110;
        for (int col=startColumn; col<endColumn; col++)
        {
            var cellValue = workSheet.Cell(row,col).Value.ToString();
            //iterates through all the employee names that we got from get shift times method
            for (int i=0; i< employees.Count; i++)
            {
                //init employee name, clock in time, and clock out time
                string employeeName = employees[i];
                string clockOutTime = clockOuts[i];
                string clockInTime = clockIns[i];
                if (cellValue.Contains(employeeName))
                {
                    int clockOutCol = col;
                    int clockInCol = col-3;
                    int dateRow = int.Parse(dayNum);
                    // the row where the dates are 4 more than what the date actually is
                    // day 1 is row 5, day 10 is row 14
                    dateRow +=4;
                    workSheet.Cell(row: dateRow, column: clockOutCol).Value = clockOutTime;
                    workSheet.Cell(row: dateRow, column: clockInCol).Value = clockInTime;
                }
            }
        }
    }

    public void Save()
    {
        _workbook.Save();
        // gets the filepath of the workbook which includes the name of the excel file
        string workBookName = _workbook.ToString();
        // gets the length of the string of the work book name, and -1, i don't want to print the last ) of the string
        int workBookNameLength = workBookName.Length-1;
        Console.WriteLine($"{workBookName[48..workBookNameLength]} File Closed");
    }

// closes the excel file
    public void Dispose()
    {
        // gets the filepath of the workbook which includes the name of the excel file
        string workBookName = _workbook.ToString();
        // gets the length of the string of the work book name, and -1, i don't want to print the last ) of the string
        int workBookNameLength = workBookName.Length-1;
        Console.WriteLine($"{workBookName[48..workBookNameLength]} File Closed");
        _workbook?.Dispose();
    }

}
