using System;
using System.IO;


class Program
{
    static void Main(string[] args)
    {
        //Для работы необходимо добавить почту с возможностью редактировать exel таблицы (sheet).
        //Почта: "sheetservise@sheets-371512.iam.gserviceaccount.com"

        //idSheet берём из ссылки на нашу таблицу между d/.../edit : https://docs.google.com/spreadsheets/d/1KdBMWjLZJ-_Nhw3WHMOdEUUd-1Ln5tdf-zgCp6fc8D4/edit#gid=0
        string idSheet = "13KXwtNkdf1Duo5bYCPrEKqb6RDYnl9JMcGE8PXR4MHE";

        //nameSheet берём снизу слева (обычно "Лист 1") как на картинке https://i.imgur.com/lmJdBmC.png 
        string nameSheet = "Sheet1";

        //nameProgect уникальное имя проекта на каждую таблицу свое уникальное
        string nameProgect = "Currenew1ed3 Legislators";

        //создаем экземпдяр SheetHelper
        Sheet.SheetHelper program = new Sheet.SheetHelper();

        //устанавливаем настройки 
        program.SetProperty(idSheet, nameSheet, "C:\\Users\\User\\source\\repos\\ConsoleApp2\\ConsoleApp2\\sheetService.json", nameProgect);
        
        //Читаем строку 
        program.PrintEntries(program.ReadEntries("a1","b32"));
        Console.WriteLine(program.ReadEntry("A1"));

       

    }
}


