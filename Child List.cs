/*
 Здесь выводить лист детей, а так-же делается excel выгрузка по детям. Здесь выполнена оптимизация в первую очередь. 
Очень медленно работал поиск по именам. Были доратки, например множественный выбор смен, многое честно уже не помню
 */

//До
namespace RestChild.Web.Controllers
{
    /// <summary>
    ///     реестр детей
    /// </summary>
    public partial class NewBoutController
    {
        /// <summary>
        ///     подготовить модель
        /// </summary>
        private IQueryable<Account> PrepareChildListModel(ChildListModel model)
        {
            if (model.PageNumber <= 0)
            {
                model.PageNumber = 1;
            }

            model.Times = MobileUw.GetSet<GroupedTime>().OrderBy(g => g.Id).ToArray();
            model.Camp = MobileUw.GetById<Camp>(model.CampId);

            var query = MobileUw.GetSet<Account>().Where(b =>
                b.IsDeleted != true && b.IsBlocked != true);

            query = query.Where(q => q.Campers.Any());

            var bouts = MobileUw.GetSet<Bout>().AsQueryable();

            var boutPresent = false;

            if (model.CampId.HasValue)
            {
                boutPresent = true;
                bouts = bouts.Where(q => q.CampId == model.CampId);
            }

            if (model.GroupedTime.HasValue)
            {
                boutPresent = true;
                bouts = bouts.Where(q => q.GroupedTimeId == model.GroupedTime);
            }

            if (!string.IsNullOrWhiteSpace(model.City))
            {
                boutPresent = true;
                var city = model.City.ToLower().Trim();
                bouts = bouts.Where(q => q.Camp.NearestCity.ToLower().Contains(city));
            }

            if (!string.IsNullOrWhiteSpace(model.Name))
            {
                var t = model.Name.ToLower().Trim();
                query = query.Where(q =>
                    q.Name.ToLower().Contains(t) || q.Campers.Any(c => c.Name.ToLower().Contains(t)));
            }

            if (boutPresent)
            {
                query = query.Where(q => q.Campers.Any(c => bouts.Any(b => b.Id == c.BoutId)));
            }

            return query;
        }

        /// <summary>
        ///     список
        /// </summary>
        public ActionResult ChildList(ChildListModel model)
        {
            if (!Security.HasRight(AccessRightEnum.NewBout.View))
            {
                return RedirectToAvailableAction();
            }

            model = model ?? new ChildListModel();

            var query = PrepareChildListModel(model);

            var pageSize = Settings.Default.TablePageSize;
            var pageNumber = model.PageNumber;
            var startRecord = (pageNumber - 1) * pageSize;
            var totalCount = query.Count();
            var items = query
                .OrderBy(b => b.Name)
                .ThenBy(b => b.Id).Skip(startRecord).Take(pageSize)
                .ToList();
            model.Result = new CommonPagedList<Account>(items, pageNumber, pageSize, totalCount);

            return View(model);
        }

        /// <summary>
        ///     выгрузка
        /// </summary>
        public ActionResult ExcelChildList(ChildListModel model)
        {
            if (!Security.HasRight(AccessRightEnum.NewBout.View))
            {
                return RedirectToAvailableAction();
            }

            var query = PrepareChildListModel(model ?? new ChildListModel());

            var columns = new List<ExcelColumn<Account>>
            {
                new ExcelColumn<Account> {Title = "Псевдоним", Func = t => t.Name ?? "-", Width = 42},
                new ExcelColumn<Account>
                {
                    Title = "ФИО", Func = t => string.Join("\n", t.Campers.OrderByDescending(c => c.Id).Select(c=>c.Name).Distinct()),
                    Width = 35
                },
                new ExcelColumn<Account>
                    {Title = "E-mail", Func = t => t.Email, Width = 30},
                new ExcelColumn<Account>
                    {Title = "Телефон", Func = t => t.Phone, Width = 15},
                new ExcelColumn<Account>
                    {Title = "Заданий выполнено", Func = t => t.TaskCount, Width = 10},
                new ExcelColumn<Account>
                    {Title = "Сумма баллов всего", Func = t => t.Points.ToString("0"), Width = 10},
                new ExcelColumn<Account>
                    {Title = "Баллов можно потратить", Func = t => t.PointsOnAccount.ToString("0"), Width = 10}
            };

            columns = columns.Select(c =>
            {
                c.WordWrap = true;
                c.VerticalAlignment = ExcelVerticalAlignment.Center;
                return c;
            }).ToList();


            var data = query.OrderBy(b => b.Name)
                .ThenBy(b => b.Id).ToList();

            using (var excel = new ExcelTable<Account>(columns))
            {
                const int startRow = 1;
                var excelWorksheet = excel.CreateExcelWorksheet("Реестр статистики по детям");

                excel.TableName = "Реестр статистики по детям";

                excel.Parameters = new List<Tuple<string, string>>
                    {
                        new Tuple<string, string>("ФИО:", model?.Name),
                        new Tuple<string, string>("Место отдыха:", model.Camp?.Name),
                        new Tuple<string, string>("Смена:", model.Times?.Where(i => i.Id == model?.GroupedTime).Select(x=>x.Name).FirstOrDefault()),
                        new Tuple<string, string>("Город:", model?.City)
                    }
                .Where(i => !String.IsNullOrWhiteSpace(i.Item2))
                .ToList();

                excel.DataBind(excelWorksheet, data, ExcelBorderStyle.Thin, startRow);

                return File(excel.CreateExcel(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Отдыхающие.xlsx");
            }
        }


        /// <summary>
        ///     Карточка для управления
        /// </summary>
        public ActionResult ManageChild(long id, string activeTab)
        {
            if (!Security.HasRight(AccessRightEnum.NewBout.View))
            {
                return RedirectToAction("Index", "Home");
            }

            var entity = MobileUw.GetById<Account>(id);
            if (entity == null)
            {
                return RedirectToAction("List");
            }

            var model = new ChildModel(entity) { ActiveTab = activeTab };

            return View(model);
        }
    }
}

//После
namespace RestChild.Web.Controllers
{
    /// <summary>
    ///     реестр детей
    /// </summary>
    public partial class NewBoutController
    {

        /// <summary>
        ///     подготовить модель
        /// </summary>
        private IQueryable<ChildListResponseModel> PrepareChildListModel(ChildListModel model)
        {
            //подготавливаем моедль фильтров
            model.PageNumber = model.PageNumber > 0 ? model.PageNumber : 1;
            model.Times = MobileUw.GetSet<GroupedTime>().OrderBy(g => g.Id).ToArray();
            model.Camp = MobileUw.GetById<Camp>(model.CampId);
            //проверяем даты
            var dateFrom = model.DateFrom ?? DateTime.MinValue;
            var dateTo = model.DateTo?.AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999) ?? DateTime.MaxValue; //дату "по" накидываем часы минуты и сек чтобы все задачи за эту дату попали

            string mName = !string.IsNullOrWhiteSpace(model.Name) ? model.Name.Trim().ToLower() : null
                , mCity = !string.IsNullOrWhiteSpace(model.City) ? model.City.Trim().ToLower() : null;
            IEnumerable<long?> mGroupedTimes = model.GroupedTimes != null ? model.GroupedTimes.Select(s => (long?)long.Parse(s)).ToArray() : new long?[] { };

            var optimizedQuerry = MobileUw.GetSet<Camper>()
                .Where(c => c.AccountId != null) //отсекаем пустых камперов
                .Select(c => new
                {
                    AccIsBlocked = c.Account.IsBlocked ?? false,
                    AccIsDel = c.Account.IsDeleted ?? false,
                    AccId = c.AccountId,
                    AccName = c.Account.Name,
                    AccPoints = c.Account.Campers.Select(innerCampers => innerCampers.Tasks.Where(t => (t.StateId == 4 || t.StateId == 102) && t.CompliteDate != null).Select(t => (int?)Math.Round(t.Price, 0)).Sum() ?? 0).Sum(), //4, 102 выполненные задачи
                    AccTasks = c.Account.Campers.Select(innerCampers => innerCampers.Tasks.Where(t => (t.StateId == 4 || t.StateId == 102) && t.CompliteDate != null).Count()).Sum(),                                               //4, 102 выполненные задачи
                    AccPointsLeft = c.Account.PointsOnAccount,
                    Email = c.Account.Email,
                    Phone = c.Account.Phone,
                    c.Bout.GroupedTimeId,
                    CamperName = c.Name,
                    CamperTasks = c.Tasks
                                    .Where(t => t.CompliteDate != null)
                                    .Where(t => (t.StateId == 4 || t.StateId == 102) //4, 102 - выполненные
                                        && (dateFrom <= t.CompliteDate) && (dateTo >= t.CompliteDate)) //фильтр по дате 
                                    .Count(),
                    CamperPoints = c.Tasks
                                    .Where(t => t.CompliteDate != null)
                                    .Where(t => (t.StateId == 4 || t.StateId == 102)
                                        && (dateFrom <= t.CompliteDate) && (dateTo >= t.CompliteDate)) //фильтр по дате 
                                    .Select(t => (int?)Math.Round(t.Price, 0)).Sum() ?? 0,
                    c.Bout.CampId,
                    c.Bout.DateIncome,
                    c.Bout.DateOutcome,
                    c.Bout.Camp.NearestCity,
                    CampName = c.Bout.Camp.Name
                }) //собираем модель с нужными нам данными
                .Where(row => !row.AccIsBlocked & !row.AccIsDel & !(row.GroupedTimeId == null)) //отбрасываем мусор
                .Where(row => //применяем фильтры, кроме имени
                    (mGroupedTimes.Count() > 0 ? mGroupedTimes.Any(gr => gr == row.GroupedTimeId) : true) //фильтр групповой по сменам
                                                                                                          //& ((((model.DateFrom ?? DateTime.MinValue)<=row.DateIncome) & ((model.DateTo ?? DateTime.MaxValue)>=row.DateIncome)) //фильтр по дате заезда
                                                                                                          //    | ((model.DateFrom ?? DateTime.MinValue)>= row.DateOutcome) & ((model.DateTo ?? DateTime.MaxValue) <= row.DateOutcome)) //фильтр по дате заезда
                    & (model.CampId != null ? row.CampId == model.CampId : true) //фильтр по лагерю
                    & (mCity != null ? row.NearestCity.ToLower().Contains(mCity) : true)) //фильтр по ближ. городу
                .GroupBy(g => new { g.AccId, g.AccName, g.AccPoints, g.AccTasks, g.AccPointsLeft, g.Email, g.Phone })
                .Select(acc => new //собираем модель, чтобы выбрать имена
                {
                    Id = (long)acc.Key.AccId,
                    AccountName = acc.Key.AccName,
                    Email = acc.Key.Email,
                    Phone = acc.Key.Phone,
                    TotalTasksCount = acc.Key.AccTasks,
                    TotalPointsCount = (int)acc.Key.AccPoints,
                    TotalPointsLeft = (int)acc.Key.AccPointsLeft,
                    Name = acc.GroupBy(n => n.CamperName).Select(s => s.Key),
                    PlaceOfRest = acc.Select(p => new ChildListResponseModel.PlaceOfRestModel { Id = p.CampId ?? -1, Name = p.CampName, Points = p.CamperPoints, Tasks = p.CamperTasks }),
                    FilteredPointsCount = (int)acc.Sum(p => p.CamperPoints),
                    FilteredTasksCount = acc.Sum(p => p.CamperTasks)
                })
                .Where(m => (mName != null ?
                    (m.AccountName.Trim().ToLower().Contains(mName)
                    | m.Name.Any(n => n.Trim().ToLower().Contains(mName)))
                    : true)) //делаем выборку по именам
                .Select(a => new ChildListResponseModel
                {
                    Id = a.Id,
                    AccountName = a.AccountName,
                    Email = a.Email,
                    Phone = a.Phone,
                    TotalTasksCount = a.TotalTasksCount,
                    TotalPointsCount = a.TotalPointsCount,
                    TotalPointsLeft = a.TotalPointsLeft,
                    Name = a.Name,
                    PlaceOfRests = a.PlaceOfRest,
                    FilteredPointsCount = a.FilteredPointsCount,
                    FilteredTasksCount = a.FilteredTasksCount
                });

            return optimizedQuerry;
        }

        /// <summary>
        ///     список
        /// </summary>
        public ActionResult ChildList(ChildListModel model)
        {
            if (model.GroupedTimesPagerArgs != null)
            {
                model.GroupedTimes = model.GroupedTimesPagerArgs.Split(',');
            }

            if (!Security.HasRight(AccessRightEnum.NewBout.View))
            {
                return RedirectToAvailableAction();
            }

            model = model ?? new ChildListModel();

            var query = PrepareChildListModel(model);

            var pageSize = Settings.Default.TablePageSize;
            var pageNumber = model.PageNumber;
            var startRecord = (pageNumber - 1) * pageSize;
            var totalCount = query.Count();
            var items = query
                .OrderBy(b => b.AccountName)
                .ThenBy(b => b.Id)
                .Skip(startRecord)
                .Take(pageSize)
                .ToList();
            model.Result = new CommonPagedList<ChildListResponseModel>(items, pageNumber, pageSize, totalCount);

            return View(model);
        }

        /// <summary>
        ///     выгрузка
        /// </summary>
        public ActionResult ExcelChildList(ChildListModel model)
        {
            if (model.GroupedTimesPagerArgs != null)
            {
                model.GroupedTimes = model.GroupedTimesPagerArgs.Split(',');
            }

            if (!Security.HasRight(AccessRightEnum.NewBout.View))
            {
                return RedirectToAvailableAction();
            }

            var query = PrepareChildListModel(model ?? new ChildListModel());

            var columns = new List<ExcelColumn<ChildListResponseModel>>
            {
                new ExcelColumn<ChildListResponseModel> {Title = "Псевдоним", Func = t => t.AccountName ?? "-", Width = 42},
                new ExcelColumn<ChildListResponseModel>
                {
                    Title = "ФИО", Func = t => string.Join("\n", t.Name.Cast<string>()),
                    Width = 35
                },
                //new ExcelColumn<ChildListResponseModel>
                //{
                //    Title = "Места отдыха", Func = t => string.Join("\n", t.PlaceOfRests.Select(p => p.Name + ":\n ● Баллы: " + p.Points + ",\n ● Задания: " + p.Tasks + ";")),
                //    Width = 35
                //},
                new ExcelColumn<ChildListResponseModel>
                    {Title = "E-mail", Func = t => t.Email, Width = 30},
                new ExcelColumn<ChildListResponseModel>
                    {Title = "Телефон", Func = t => t.Phone, Width = 15},
                new ExcelColumn<ChildListResponseModel>
                    {Title = "Заданий выполнено", Func = t => t.TotalTasksCount, Width = 10},
                new ExcelColumn<ChildListResponseModel>
                    {Title = "Сумма баллов всего", Func = t => t.TotalPointsCount.ToString(), Width = 10},
                new ExcelColumn<ChildListResponseModel>
                    {Title = "Баллов можно потратить", Func = t => t.TotalPointsLeft.ToString(), Width = 10},
                new ExcelColumn<ChildListResponseModel>
                    {Title = "Выполненные задания за период", Func = t => t.FilteredTasksCount.ToString(), Width = 10},
                new ExcelColumn<ChildListResponseModel>
                    {Title = "Сумма баллов за период", Func = t => t.FilteredPointsCount.ToString(), Width = 10}
            };

            columns = columns.Select(c =>
            {
                c.WordWrap = true;
                c.VerticalAlignment = ExcelVerticalAlignment.Center;
                return c;
            }).ToList();


            var data = query.OrderBy(b => b.AccountName)
                .ThenBy(b => b.Id).ToList();

            using (var excel = new ExcelTable<ChildListResponseModel>(columns))
            {
                const int startRow = 1;
                var excelWorksheet = excel.CreateExcelWorksheet("Реестр статистики по детям");

                excel.TableName = "Реестр статистики по детям";

                excel.Parameters = new List<Tuple<string, string>>();
                if (model.DateFrom.HasValue || model.DateTo.HasValue)
                    excel.Parameters.Add(new Tuple<string, string>("Дата:", "c " + (model.DateFrom?.ToString("dd.MM.yyyy") ?? "∞") + " по " + (model.DateTo?.ToString("dd.MM.yyyy") ?? "∞\n")));
                if (model.GroupedTimes != null || model.GroupedTimes?.Length > 0)
                    excel.Parameters.Add(new Tuple<string, string>("Смены:", String.Join(", ", model.GroupedTimes.Select(t => ToRomanNumber(Convert.ToInt32(t))))));
                if (!string.IsNullOrWhiteSpace(model.Name))
                    excel.Parameters.Add(new Tuple<string, string>("ФИО:", model.Name));
                if (model.CampId.HasValue)
                    excel.Parameters.Add(new Tuple<string, string>("Место отдыха:", model.Camp.Name));
                if (!string.IsNullOrWhiteSpace(model.City))
                    excel.Parameters.Add(new Tuple<string, string>("Город:", model.City));


                excel.DataBind(excelWorksheet, data, ExcelBorderStyle.Thin, startRow);

                return File(excel.CreateExcel(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Отдыхающие.xlsx");
            }
        }

        /// <summary>
        /// Перевод Арабских чисел на Римские
        /// </summary>
        /// <param name="number">арабское число (от 0 до 3999)</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        private string ToRomanNumber(int number)
        {
            if ((number < 0) || (number > 3999)) return number.ToString();
            if (number < 1) return string.Empty;
            if (number >= 1000) return "M" + ToRomanNumber(number - 1000);
            if (number >= 900) return "CM" + ToRomanNumber(number - 900);
            if (number >= 500) return "D" + ToRomanNumber(number - 500);
            if (number >= 400) return "CD" + ToRomanNumber(number - 400);
            if (number >= 100) return "C" + ToRomanNumber(number - 100);
            if (number >= 90) return "XC" + ToRomanNumber(number - 90);
            if (number >= 50) return "L" + ToRomanNumber(number - 50);
            if (number >= 40) return "XL" + ToRomanNumber(number - 40);
            if (number >= 10) return "X" + ToRomanNumber(number - 10);
            if (number >= 9) return "IX" + ToRomanNumber(number - 9);
            if (number >= 5) return "V" + ToRomanNumber(number - 5);
            if (number >= 4) return "IV" + ToRomanNumber(number - 4);
            if (number >= 1) return "I" + ToRomanNumber(number - 1);
            throw new ArgumentOutOfRangeException("something bad happened");
        }

        /// <summary>
        ///     Карточка для управления
        /// </summary>
        public ActionResult ManageChild(long id, string activeTab)
        {
            if (!Security.HasRight(AccessRightEnum.NewBout.View))
            {
                return RedirectToAction("Index", "Home");
            }

            var entity = MobileUw.GetById<Account>(id);
            if (entity == null)
            {
                return RedirectToAction("List");
            }

            var model = new ChildModel(entity) { ActiveTab = activeTab };

            return View(model);
        }
    }
}