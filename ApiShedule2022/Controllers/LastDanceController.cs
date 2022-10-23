using ApiShedule2022.Models;
using ClosedXML.Excel;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace ApiShedule2022.Controllers
{
    [Controller]
    [Route("api/{controller}")]
    public class LastDanceController : ControllerBase
    {
        IMemoryCache cache;
        public LastDanceController(IMemoryCache cache)
        {
            this.cache = cache;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ru-RU");
        }

        [HttpGet("getnes")]
        public ActionResult<string> GetNew() //Метод для возвращения информации о наличии нового расписания
        {
            IXLWorksheet result = null;
            if (cache.TryGetValue<IXLWorksheet>("xLNew", out result))
            {
                if (result == null)
                {
                    return BadRequest("нет нового расписания");
                }
                else
                {
                    return Ok("есть новое расписание");
                }
            }
            else
            {
                return BadRequest("нет нового расписания");
            }
        }

        [HttpGet("getteacher/{teacher}")]
        public ActionResult<IEnumerable<List<DayWeekClass>>> GetTeach(string teacher, DateTime date) //Метод для возвращения расписания по преподавателям
        {
            IXLWorksheet result = null;
            List<DayWeekClass> days = new List<DayWeekClass>();
            while (result is null)
            {
                if (DateTime.UtcNow.AddHours(5).Date == date.Date)
                {
                    result = (IXLWorksheet)cache.Get("xLMain");
                    days.Add(new DayWeekClass { Day = "ЧКР" });
                    for (int i = 1; i <= 6; i++)
                    {
                        int row = 6;
                        List<DayWeekClass> teachers = Differents.raspisanieteach(row * i, teacher, result);
                        days.AddRange(teachers.ToArray());
                    }
                }
                else if (DateTime.UtcNow.AddMinutes(5).Date.AddDays(7) == date.Date)
                {
                    result = (IXLWorksheet)cache.Get("xLNew");
                    days.Add(new DayWeekClass { Day = "ЧКР" });
                    for (int i = 1; i <= 6; i++)
                    {
                        int row = 6;
                        List<DayWeekClass> teachers = Differents.raspisanieteach(row * i, teacher, result);
                        days.AddRange(teachers.ToArray());
                    }
                }
                else
                {
                    result = Differents.dictSpecial[date.Date].Item1;
                    days.Add(new DayWeekClass { Day = "ЧКР" });
                    for (int i = 1; i <= 6; i++)
                    {
                        int row = 6;
                        List<DayWeekClass> teachers = Differents.raspisanieteach(row * i, teacher, result);
                        days.AddRange(teachers.ToArray());
                    }
                }
            }
            return Ok(days);
        }

        [HttpGet("getgroup/{group}")]
        public ActionResult<IEnumerable<List<DayWeekClass>>> Get(string group, DateTime date) //Метод для возвращения расписания по группам
        {
            IXLWorksheet result = null;
            List<DayWeekClass> days = new List<DayWeekClass>();
            while (result is null)
            {
                if (DateTime.UtcNow.AddHours(5).Date == date.Date)
                {
                    result = (IXLWorksheet)cache.Get("xLMain");
                    days.Add(new DayWeekClass { Day = "ЧКР" });
                    int column = Differents.IndexGroup(group, result);
                    for (int j = 1; j <= 6; j++)
                    {
                        List<DayWeekClass> metrics = Differents.EnumerableMetrics(j * 6, column, result);
                        days.AddRange(metrics.ToArray());
                    }
                }
                else if (DateTime.UtcNow.AddMinutes(5).Date.AddDays(7) == date.Date)
                {
                    result = (IXLWorksheet)cache.Get("xLNew");
                    days.Add(new DayWeekClass { Day = "ЧКР" });
                    int column = Differents.IndexGroup(group, result);
                    for (int j = 1; j <= 6; j++)
                    {
                        List<DayWeekClass> metrics = Differents.EnumerableMetrics(j * 6, column, result);
                        days.AddRange(metrics.ToArray());
                    }
                }
                else
                {
                    result = Differents.dictSpecial[date.Date].Item1;
                    days.Add(new DayWeekClass { Day = "ЧКР" });
                    int column = Differents.IndexGroup(group, result);
                    for (int j = 1; j <= 6; j++)
                    {
                        List<DayWeekClass> metrics = Differents.EnumerableMetrics(j * 6, column, result);
                        days.AddRange(metrics.ToArray());
                    }
                }
            }
            return Ok(days);
        }

        [HttpGet("getcabinet/{сabinet}")]
        public ActionResult<IEnumerable<List<DayWeekClass>>> GetKab(string сabinet, DateTime date) //Метод для возвращения расписания по кабинетам
        {
            IXLWorksheet result = null;
            List<DayWeekClass> days = new List<DayWeekClass>();
            while (result is null)
            {
                if (DateTime.UtcNow.AddHours(5).Date == date.Date)
                {
                    result = (IXLWorksheet)cache.Get("xLMain");
                    days.Add(new DayWeekClass { Day = "ЧКР" });
                    for (int i = 1; i <= 6; i++)
                    {
                        int row = 6;
                        List<DayWeekClass> сabinets = Differents.raspisaniekab(row * i, сabinet, result);
                        days.AddRange(сabinets.ToArray());
                    }
                }
                else if (DateTime.UtcNow.AddMinutes(5).Date.AddDays(7) == date.Date)
                {
                    result = (IXLWorksheet)cache.Get("xLNew");
                    days.Add(new DayWeekClass { Day = "ЧКР" });
                    for (int i = 1; i <= 6; i++)
                    {
                        int row = 6;
                        List<DayWeekClass> сabinets = Differents.raspisaniekab(row * i, сabinet, result);
                        days.AddRange(сabinets.ToArray());
                    }
                }
                else
                {
                    result = Differents.dictSpecial[date.Date].Item1;
                    days.Add(new DayWeekClass { Day = "ЧКР" });
                    for (int i = 1; i <= 6; i++)
                    {
                        int row = 6;
                        List<DayWeekClass> сabinets = Differents.raspisaniekab(row * i, сabinet, result);
                        days.AddRange(сabinets.ToArray());
                    }
                }
            }
            return Ok(days);
        }

        [HttpGet("getdate/{date}")]
        public ActionResult<ListPack> GetDate(DateTime date) //Метод для возвращения расписания по конкретной неделе
        {
            ListPack listPack = new ListPack();
            if (DateTime.UtcNow.AddHours(5).Date == date)
            {
                while (listPack.Cabinets.Count() is 0 || listPack.Teachers.Count() is 0 || listPack.Groups.Count() is 0)
                {
                    listPack.Cabinets = (List<string>)cache.Get("MainListCabinets");
                    listPack.Teachers = (List<string>)cache.Get("MainListTeachers");
                    listPack.Groups = (List<string>)cache.Get("MainListGroups");
                }
            }
            else if (DateTime.UtcNow.AddHours(5).Date.AddDays(7) == date)
            {
                while (listPack.Cabinets.Count() is 0 || listPack.Teachers.Count() is 0 || listPack.Groups.Count() is 0)
                {
                    listPack.Cabinets = (List<string>)cache.Get("NewListCabinets");
                    listPack.Teachers = (List<string>)cache.Get("NewListTeachers");
                    listPack.Groups = (List<string>)cache.Get("NewListGroups");
                }
            }
            else
            {
                var lp = Differents.SpecialSheduleReturn(date);
                if (lp == null)
                    return Ok(null);
                while (listPack.Cabinets.Count() is 0 || listPack.Teachers.Count() is 0 || listPack.Groups.Count() is 0)
                {
                    listPack.Cabinets = lp.Cabinets;
                    listPack.Teachers = lp.Teachers;
                    listPack.Groups = lp.Groups;
                }
            }
            return Ok(listPack);
        }

        public static void RaspisanieIzm(XLWorkbook _workbook1, int h, IXLWorksheet ix) //Метод для применения скачанных изменений в основное расписание
        {
            DateTime dateIZM = DateTime.Today;
            int dayWeek = (int)DateTime.Today.DayOfWeek;
            var worksheet = _workbook1.Worksheets.First();
            for (int i = 1; i <= worksheet.ColumnsUsed().Count(); i++)
            {
                int n = worksheet.RowsUsed().Count();
                for (int j = 11; j <= worksheet.RowsUsed().Count() + 10; j++)
                {
                    for (int l = 3; l <= ix.ColumnsUsed().Count(); l++)
                    {
                        if (ix.Cell(5, l).GetValue<string>() == worksheet.Cell(j, i).GetValue<string>())
                        {
                            bool a = false;
                            int g = 6;
                            for (int m = 1; m <= g; m++)
                            {
                                IXLCell leg = worksheet.Cell(j + m, i);
                                if (leg.Style.Font.FontSize >= 22 || leg.Value.ToString() == "" || a || leg.Value.ToString().Length == 4)
                                {
                                    if (ix.Cell(27, 2).Value.ToString() != "4")
                                        ix.Cell((6 * h) + m, l).Value = " ";
                                    else
                                        ix.Cell((6 * h) + m - 1, l).Value = " ";
                                    a = true;
                                }
                                else
                                {
                                    if (ix.Cell(27, 2).Value.ToString() != "4")
                                        ix.Cell((6 * h) + m, l).Value = worksheet.Cell(j + m, i);
                                    else
                                        ix.Cell((6 * h) + m - 1, l).Value = worksheet.Cell(j + m, i);
                                }
                            }
                        }
                    }
                }
            }
        }

        [HttpGet]
        [Route("getnewWeek/{btnContent}")]
        public async Task<ActionResult<ListPack>> GetNew(string btnContent) //Метод для получения нового или старого расписания в зависимости от тела запроса
        {
            if (btnContent == "Новое расписание!!!")
            {
                return GetDate(DateTime.UtcNow.AddHours(5).AddDays(7).Date);
            }
            else
            {
                return GetDate(DateTime.UtcNow.AddHours(5));
            }
        }

        [HttpGet]
        [Route("getSignData/{data}")]
        public async Task<ActionResult> GetSign(string data) //Метод для выполнения первого входа в мобильное приложение с помощью пароля
        {
            if (data == "Mat'NeTrogai")
                return Ok();
            else
                return BadRequest();
        }

        [HttpGet]
        [Route("searchEmptycabinet/{numberPara}")]
        public async Task<ActionResult> SearchEmptyCabinet(int numberPara)
        {
            var fullShed = (IXLWorksheet)cache.Get("xLMain");
            int actualRow = (int)DateTime.UtcNow.AddHours(5).DayOfWeek*6+numberPara-1;
            int columnsCount = fullShed.ColumnsUsed().Count();
            var cabinets = (List<string>)cache.Get("MainListCabinets");
            var emptyCabinetsNow = new List<string>();
            foreach (var cabinet in cabinets)
            {
                bool cont = false;
                for (int i = 3; i <= columnsCount; i++)
                {
                    var cell = fullShed.Cell(actualRow, i).GetValue<string>().Trim();
                    if (cell.Contains(cabinet))
                    {
                        cont = true;
                        break;
                    }
                }
                if (cont)
                {
                    cont = false;
                    continue;
                }
                emptyCabinetsNow.Add(cabinet);
            }
            return Ok(emptyCabinetsNow);
        }
    }
}
