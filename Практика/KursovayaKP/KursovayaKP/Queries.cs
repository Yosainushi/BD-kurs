using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KursovayaKP
{
    class Queries
    {
        public static string selectZakazi = "SELECT Заказы.Id_Заказа, Заказчики.Фамилия +' '+ Заказчики.Имя + ' ' + Заказчики.Отчество AS [ФИО Заказчика], Работники.Фамилия + ' ' + Работники.Имя + ' ' + Работники.Отчество AS [оформляющий], Заказы.Адрес_Помещения, Заказы.Сумма_Заказа, Заказы.Дата_Заказа FROM Работники INNER JOIN(Заказчики INNER JOIN Заказы ON Заказчики.Id_Заказчика = Заказы.Id_Заказчика) ON Работники.Id_Работника = Заказы.Оформляющий;";
        public static string selectZakazch = "SELECT * FROM Заказчики;";
        public static string selectMat = "SELECT Материалы.Код_Материала, Материалы.Наименование, Материалы.Цена_За_1, Материалы.Колво, Склады.Адрес_Склада FROM Склады INNER JOIN Материалы ON Склады.id_Склада = Материалы.Id_Склада;";
        public static string selectPolz = "SELECT * FROM Пользователи;";
        public static string selectPostavki = "SELECT Поставки.Id_Поставки, Поставщики.Название_организации, Материалы.Наименование, Склады.Адрес_Склада, Поставки.Дата_Разгрузки, Работники.Фамилия FROM Работники INNER JOIN(Поставщики INNER JOIN (Материалы INNER JOIN (Склады INNER JOIN Поставки ON Склады.id_Склада = Поставки.Id_Склада) ON(Склады.id_Склада = Материалы.Id_Склада) AND(Материалы.Код_Материала = Поставки.Id_Материала)) ON Поставщики.Id_поставщика = Поставки.Id_Поставщика) ON Работники.Id_Работника = Поставки.Ответсвенный_За_Поставку;";
        public static string selectPostavch = "SELECT * FROM Поставщики;";
        public static string selectPrikazi = "SELECT Приказы.Id_Приказа, Приказы.Id_Состава_Заказа, Приказы.Список_Работников FROM Приказы INNER JOIN Состав_Заказа ON Приказы.Id_Состава_Заказа = Состав_Заказа.Id_Состава_Заказа;";
        public static string selectRabotniki = "SELECT * FROM Работники;";
        public static string selectScladi = "SELECT * FROM Склады;";

    }
}
