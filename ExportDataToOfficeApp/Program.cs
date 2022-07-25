using System;
using System.Collections.Generic;
// Создать псевдоним для объектной модели Excel.
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportDataToOfficeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Car> carsInStock = new List<Car>
            {
                new Car{ Color = "Green", Make = "VW", PetName = "Mary" },
                new Car{ Color = "Red", Make = "Saab", PetName = "Mel" },
                new Car{ Color = "Black", Make = "Ford", PetName = "Hank" },
                new Car{ Color = "Yellow", Make = "BMW", PetName = "Davie" }
            };
            ExportToExcel(carsInStock);
        }
        static void ExportToExcel(List<Car> carInStock)
        {
            // Загрузить Excel и затем создать новую пустую рабочую книгу.
            Excel.Application excelApp = new Excel.Application();
            // Сделать пользовательский интерфейс Excel видимым на рабочем столе.
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = excelApp.ActiveSheet;
            // Установить заголовки столбцов в ячейках.
            workSheet.Cells[1, "A"] = "Make";
            workSheet.Cells[1, "B"] = "Color";
            workSheet.Cells[1, "C"] = "Pet Name";
            // Отобразить все данные из List<Car> на ячейки электронной таблицы.
            int row = 1;
            foreach(Car c in carInStock)
            {
                row++;
                workSheet.Cells[row, "A"] = c.Make;
                workSheet.Cells[row, "B"] = c.Color;
                workSheet.Cells[row, "C"] = c.PetName;
            }
            // Придать смпатичный вид табличным данным.
            workSheet.Range["A1"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
            // Сохранить файл, завершить работу Excel и отобразить сообщение пользователю.
            workSheet.SaveAs2($@"{Environment.CurrentDirectory}\Inventory.xlsx");
            excelApp.Quit();
            Console.WriteLine("The inventory.xlsx file has been saved to your app filder");
        }
    }
}
