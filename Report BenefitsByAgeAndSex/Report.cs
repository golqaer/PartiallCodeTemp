/*
 Здесь выполнена оптимизация кода, ибо предыдущий запрос работал кастарофически долго. Пришллось очень плотно изучать много аспектов
бизнес логики. Этот отчет делался сначало через SQL дял упрощения
 */

//Было
private AnalyticReportInfo<FamilyTimeOfRestInfo> GetFamilyRestData(AnalyticReportFilter filter)
{
    var requests = UnitOfWork.GetSet<Request>().AsQueryable();
    var attendants = UnitOfWork.GetSet<Applicant>().AsQueryable();
    var applicants = UnitOfWork.GetSet<Applicant>().AsQueryable().Where(i => i.IsAccomp);
    var typeOfRests = UnitOfWork.GetSet<TypeOfRest>().AsQueryable();

    var query = UnitOfWork.GetSet<Child>().AsQueryable()
        .Include(i => i.Request.Hotels)
        .Include(i => i.Request.TimeOfRest)
        .Include(i => i.Request.TypeOfRest)
        .Where(i => !i.IsDeleted)
        .Where(i => i.Request != null && !i.Request.IsDeleted)
        .Where(i => i.Request.Tour != null)
        .Where(i => i.Request.StatusId == (int)StatusEnum.CertificateIssued)
        .Where(i => i.Request.Tour.TypeOfRest.ParentId == (long)TypeOfRestEnum.RestWithParents || i.Request.Tour.TypeOfRestId == (long)TypeOfRestEnum.RestWithParents);

    #region исправление подсчёта сопровождабщих через нормальный LINQ

    var naRequests = UnitOfWork.GetSet<Request>().Where(nar => nar.YearOfRestId == filter.YearOfRestId
                                                                && !nar.IsDeleted
                                                                && nar.StatusId == (long)StatusEnum.CertificateIssued
                                                                && !nar.DeclineReasonId.HasValue
                                                                && nar.Tour.TypeOfRest.ParentId == (long)TypeOfRestEnum.RestWithParents || nar.Tour.TypeOfRestId == (long)TypeOfRestEnum.RestWithParents);

    if (filter.YearOfBirthDateBegin.HasValue)
        naRequests = naRequests.Where(nar => nar.Child.Any(c => c.DateOfBirth.Value.Year >= filter.YearOfBirthDateBegin));
    if (filter.YearOfBirthDateEnd.HasValue)
        naRequests = naRequests.Where(nar => nar.Child.Any(c => c.DateOfBirth.Value.Year <= filter.YearOfBirthDateEnd));


    naRequests = naRequests.Where(nar => nar.TourId.HasValue && nar.HotelsId.HasValue);
    if (filter.HotelId.HasValue)
        naRequests = naRequests.Where(nar => nar.Hotels.Id == filter.HotelId /*&& (nar.Tour.DateIncome >= filter.DateStartBegin && nar.Tour.DateIncome <= filter.DateStartEnd)*/);
    if (filter.DateStartBegin.HasValue)
        naRequests = naRequests.Where(nar => nar.TimeOfRest.Month >= filter.DateStartBegin.Value.Month
                                && nar.TimeOfRest.DayOfMonth >= filter.DateStartBegin.Value.Day);

    if (filter.DateStartEnd.HasValue)
        naRequests = naRequests.Where(nar => nar.TimeOfRest.Month <= filter.DateStartEnd.Value.Month
                                && nar.TimeOfRest.DayOfMonth <= filter.DateStartEnd.Value.Day);

    int attendantFemale = 0;
    int attendantMale = 0;

    //var attendantFemale = naRequests.Count(nar => nar.Applicant.IsAccomp && (bool)!nar.Applicant.Male);
    //attendantFemale += attendants.Count(att => naRequests.Any(nar => nar.Id == att.RequestId) && (bool)!att.Male);
    //var attendantMale = naRequests.Count(nar => nar.Applicant.IsAccomp && (bool)nar.Applicant.Male);
    //attendantMale += attendants.Count(att => naRequests.Any(nar => nar.Id == att.RequestId) && (bool)att.Male);
    #endregion

    query = AddFilters(query, filter);

    var result = query.Select(i => new
    {
        hotel = i.Request.Tour.Hotels,
        timeOfRest = i.Request.Tour.TimeOfRest,
        i.Male,
        request = i.Request,
        DateOfBirth = i.DateOfBirth.Value.Year
    }).ToArray();

    var attendantsQuery = from request in requests
                          join typeOfRest in typeOfRests on request.TypeOfRestId equals typeOfRest.Id
                          join attendant in attendants on request.Id equals attendant.RequestId into attendantGr
                          from attendantQuery in attendantGr.DefaultIfEmpty()
                          join applicant in applicants on new { request.ApplicantId, IsAccomp = true } equals
                              new { ApplicantId = (long?)applicant.Id, applicant.IsAccomp } into applicantGr
                          from applicantQuery in applicantGr.DefaultIfEmpty()
                          where !request.IsDeleted && request.TourId.HasValue
                          where (applicantQuery == null || !applicantQuery.IsDeleted)
                          where (attendantQuery == null || !attendantQuery.IsDeleted)
                          where request.StatusId == (int)StatusEnum.CertificateIssued
                          where typeOfRest.ParentId == (long)TypeOfRestEnum.RestWithParents
                          select new { request, attendantQuery, applicantQuery };

    var attendantsByTourQuery = from attendant in attendantsQuery
                                group new { attendant.applicantQuery, attendant.attendantQuery } by attendant.request.Id
        into requestGr
                                select new
                                {
                                    RequestId = requestGr.Key,
                                    attendantMaleCount = requestGr.Count(i => (i.applicantQuery != null && i.applicantQuery.Male.HasValue && i.applicantQuery.Male.Value)
                                    || (i.attendantQuery != null && i.attendantQuery.Male.HasValue && i.attendantQuery.Male.Value)
                                    || ((i.attendantQuery != null && !i.attendantQuery.Male.HasValue) || i.applicantQuery != null && !i.applicantQuery.Male.HasValue)),
                                    attendantFemaleCount = requestGr.Count(i => (i.applicantQuery != null && i.applicantQuery.Male.HasValue && !i.applicantQuery.Male.Value)
                                    || (i.attendantQuery != null && i.attendantQuery.Male.HasValue && !i.attendantQuery.Male.Value))
                                };

    var bookedApplicantsByTour = attendantsByTourQuery.ToDictionary(i => i.RequestId, j => new { j.attendantMaleCount, j.attendantFemaleCount });

    var data = new AnalyticReportInfo<FamilyTimeOfRestInfo>
    {
        SubHeaders = new Dictionary<long, string>(),
    };

    var timeOfRestInfos = new List<FamilyTimeOfRestInfo>();
    foreach (var hotel in result.GroupBy(i => i.hotel).OrderBy(i => i.Key.Id))
    {
        foreach (var timeOfRest in hotel.GroupBy(i => i.timeOfRest).OrderBy(i => i.Key.Month).ThenBy(i => i.Key.DayOfMonth))
        {
            var campTimeORestInfo = new FamilyTimeOfRestInfo();

            campTimeORestInfo.TimeOfRestName = timeOfRest.Key.Name;

            campTimeORestInfo.MaleCount = new Dictionary<long, int>();
            campTimeORestInfo.FemaleCount = new Dictionary<long, int>();

            campTimeORestInfo.HotelName = hotel.Key.Name;

            foreach (var tourKey in timeOfRest.GroupBy(i => i.request.Id))
            {
                foreach (var birthYearGr in tourKey.GroupBy(i => i.DateOfBirth))
                {
                    if (campTimeORestInfo.MaleCount.ContainsKey(birthYearGr.Key))
                        campTimeORestInfo.MaleCount[birthYearGr.Key] += birthYearGr.Count(i => i.Male);
                    else
                    {
                        campTimeORestInfo.MaleCount.Add(birthYearGr.Key, birthYearGr.Count(i => i.Male));
                    }


                    if (campTimeORestInfo.FemaleCount.ContainsKey(birthYearGr.Key))
                    {
                        campTimeORestInfo.FemaleCount[birthYearGr.Key] += birthYearGr.Count(i => !i.Male);
                    }
                    else
                    {
                        campTimeORestInfo.FemaleCount.Add(birthYearGr.Key, birthYearGr.Count(i => !i.Male));
                    }

                    if (!data.SubHeaders.ContainsKey(birthYearGr.Key))
                        data.SubHeaders.Add(new KeyValuePair<long, string>(birthYearGr.Key, birthYearGr.Key.ToString()));
                }

                if (bookedApplicantsByTour.ContainsKey(tourKey.Key))
                {
                    campTimeORestInfo.AttendantsMaleCount += bookedApplicantsByTour[tourKey.Key].attendantMaleCount;
                    campTimeORestInfo.AttendantsFemaleCount += bookedApplicantsByTour[tourKey.Key].attendantFemaleCount;
                }

            }
            // исправление через linq
            attendantFemale = naRequests.Count(nar => nar.Applicant.IsAccomp && (bool)!nar.Applicant.Male && nar.Tour.TimeOfRest.Name == timeOfRest.Key.Name && nar.Hotels.Name == hotel.Key.Name && !nar.Applicant.IsDeleted);
            attendantFemale += attendants.Count(att => naRequests.Any(nar => nar.Id == att.RequestId) && (bool)!att.Male && att.Request.Tour.TimeOfRest.Name == timeOfRest.Key.Name && att.Request.Hotels.Name == hotel.Key.Name && !att.IsDeleted);
            attendantMale = naRequests.Count(nar => nar.Applicant.IsAccomp && (bool)nar.Applicant.Male && nar.Tour.TimeOfRest.Name == timeOfRest.Key.Name && nar.Hotels.Name == hotel.Key.Name && !nar.Applicant.IsDeleted);
            attendantMale += attendants.Count(att => naRequests.Any(nar => nar.Id == att.RequestId) && (bool)att.Male && att.Request.Tour.TimeOfRest.Name == timeOfRest.Key.Name && att.Request.Hotels.Name == hotel.Key.Name && !att.IsDeleted);



            if (attendantFemale > 0)
                campTimeORestInfo.AttendantsFemaleCount = attendantFemale;
            if (attendantMale > 0)
                campTimeORestInfo.AttendantsMaleCount = attendantMale;

            timeOfRestInfos.Add(campTimeORestInfo);
        }
    }

    data.Data = timeOfRestInfos;

    return data;
}

//Стало. Скорость секунда, если не ее доли, против нескольких минут
private AnalyticReportInfo<FamilyTimeOfRestInfo> GetFamilyRestData(AnalyticReportFilter filter)
{
    var baseRequestQuery = UnitOfWork.GetSet<Request>() //базовый запрос заявлений, будет использоваться в остальных отчетах
        .Where(req => req.StatusId == (long)StatusEnum.CertificateIssued
            && req.TimeOfRestId != null
            && req.YearOfRestId == filter.YearOfRestId //фильтр года мероприятия
            && (filter.DateRestBegin != null ? req.Tour.DateIncome >= filter.DateRestBegin : true) //фильтр даты отдыха
            && (filter.DateRestEnd != null ? req.Tour.DateOutcome <= filter.DateRestEnd : true)    //фильтр даты отдыха
            && req.TypeOfRest.ParentId == 2 //2 - совместный отдых
            && req.HotelsId != null
            && req.Tour.HotelsId != null
            ).Select(r => r)
            .Where(r => (filter.HotelId != null ? r.HotelsId == filter.HotelId : true)); //фильтр по отелю, вынес отдельно условие, потому что запрос падал с null значением в HotelsId

    //подсчет детей
    var childCountStatistics = baseRequestQuery
        .Join(UnitOfWork.GetSet<Child>().Where(ch => ch.DateOfBirth != null && !ch.IsDeleted
            && (filter.YearOfBirthDateBegin != null ? ch.DateOfBirth.Value.Year >= filter.YearOfBirthDateBegin : true) //фильтр
            && (filter.YearOfBirthDateEnd != null ? ch.DateOfBirth.Value.Year <= filter.YearOfBirthDateEnd : true))   //фильтр
                , req => req.Id, ch => ch.RequestId, (req, ch) => //join детей с датами рождения и не удаленных
                    new { HotelId = req.HotelsId, HotelName = req.Hotels.Name, TimeOfRestId = req.TimeOfRestId, req.TimeOfRest.Month, TimeOfRestName = req.TimeOfRest.Name, ChildId = ch.Id, ChildMale = ch.Male, BirthYear = ch.DateOfBirth.Value.Year }) //Новая модель данных на основе join
        .GroupBy(g => new { g.HotelId, g.HotelName, g.TimeOfRestId, g.Month, g.TimeOfRestName, g.ChildMale, g.BirthYear }) //группировка по полям
        .Select(result => new { result.Key.HotelId, result.Key.HotelName, result.Key.TimeOfRestId, result.Key.Month, result.Key.TimeOfRestName, result.Key.BirthYear, result.Key.ChildMale, ChildCount = result.Count() }) //итоговый отчет по детям
        .ToArray();
    //подсчет сопроводов
    var attendantCountStatistic = baseRequestQuery
        .Join(UnitOfWork.GetSet<Applicant>().Where(att => att.Male != null && att.IsAccomp && !att.IsDeleted && att.DateOfBirth != null), req => req.ApplicantId, att => att.Id, (req, att) => //join заявителей являющиеся споровождающими, не удалены и не с пустым полом
            new { HotelId = req.HotelsId, HotelName = req.Hotels.Name, TimeOfRestId = req.TimeOfRestId, TimeOfRestName = req.TimeOfRest.Name, AttMale = att.Male, AttId = att.Id, BirthYear = att.DateOfBirth.Value.Year })
        .Union(baseRequestQuery
            .Join(UnitOfWork.GetSet<Applicant>().Where(att => att.Male != null && !att.IsDeleted && att.DateOfBirth != null), req => req.Id, att => att.RequestId, (req, att) => //join сопроводов не являющихся заявителями, не удалены и не с пустым полом
                new { HotelId = req.HotelsId, HotelName = req.Hotels.Name, TimeOfRestId = req.TimeOfRestId, TimeOfRestName = req.TimeOfRest.Name, AttMale = att.Male, AttId = att.Id, BirthYear = att.DateOfBirth.Value.Year }))
        .GroupBy(g => new { g.HotelId, g.HotelName, g.TimeOfRestId, g.TimeOfRestName, g.AttMale, g.BirthYear })
        .Select(s => new { s.Key.HotelId, s.Key.HotelName, s.Key.TimeOfRestId, s.Key.TimeOfRestName, s.Key.AttMale, s.Key.BirthYear, AttCount = s.Count() })
        .GroupBy(g => new { g.HotelId, g.HotelName, g.TimeOfRestId, g.TimeOfRestName })
        .Select(s => new { s.Key.HotelId, s.Key.HotelName, s.Key.TimeOfRestId, s.Key.TimeOfRestName, CountBySex = s.Select(cbs => new { cbs.BirthYear, cbs.AttMale, cbs.AttCount }) })
        .ToArray();

    var dataResult = new AnalyticReportInfo<FamilyTimeOfRestInfo>
    {
        SubHeaders = childCountStatistics
            .GroupBy(g => g.BirthYear)
            .ToDictionary(keySelector: k => (long)k.Key, elementSelector: e => e.Key.ToString() + " Дети") //собираем хедеры по примеру 2004...2019

    };
    dataResult.SubHeaders.Add(-1, "Всего Детей"); //добавляем подзаголовок всего под -1

    //добавляем сопроводов по годам в заголовки
    attendantCountStatistic
        .SelectMany(s => s.CountBySex).Select(s => s.BirthYear).Distinct()
        .ToDictionary(keySelector: k => (long)k, elementSelector: e => e.ToString() + " Сопровождающие")
        .ToList().ForEach(d => { dataResult.SubHeaders.Add(d.Key, d.Value); });

    //Группируем - вкладываем в отели времена отдыха, во времена разбивку по годам рождения, в года счетчик муж и жен
    dataResult.Data = childCountStatistics.GroupBy(g => new { g.HotelId, g.HotelName, g.TimeOfRestId, g.Month, g.TimeOfRestName, g.BirthYear })
        .Select(s => new { s.Key.HotelId, s.Key.HotelName, s.Key.TimeOfRestId, s.Key.Month, s.Key.TimeOfRestName, s.Key.BirthYear, CountBySex = s.Select(cbs => new { cbs.ChildMale, cbs.ChildCount }).ToArray() })
        .GroupBy(g => new { g.HotelId, g.HotelName, g.TimeOfRestId, g.Month, g.TimeOfRestName })
        .Select(s => new { s.Key.HotelId, s.Key.HotelName, s.Key.TimeOfRestId, s.Key.Month, s.Key.TimeOfRestName, CountByYear = s.Select(cby => new { cby.BirthYear, cby.CountBySex }).ToArray() })
        .Join(attendantCountStatistic, ch => new { ch.HotelId, ch.TimeOfRestId }, att => new { att.HotelId, att.TimeOfRestId }, (ch, att) => //присоединяем кол-во сопроводов
            new { ch.HotelId, ch.HotelName, ch.TimeOfRestName, ch.Month, att.CountBySex, ch.CountByYear })
        .OrderBy(o => o.HotelId).ThenBy(o => o.Month)
        .Select(st =>
            new FamilyTimeOfRestInfo()
            {
                HotelName = st.HotelName
               ,
                TimeOfRestName = st.TimeOfRestName
               ,
                AttendantsMaleCount = st.CountBySex.Where(c => (bool)c.AttMale).Select(s => s.AttCount).FirstOrDefault() //всего колонка справа
               ,
                AttendantsFemaleCount = st.CountBySex.Where(c => !(bool)c.AttMale).Select(s => s.AttCount).FirstOrDefault() //всего колонка справа
               ,
                MaleCount = st.CountByYear
                        .ToDictionary(keySelector: k => (long)k.BirthYear, elementSelector: e => (int)e.CountBySex.Where(c => (bool)c.ChildMale).Select(s => s.ChildCount).FirstOrDefault())
               ,
                FemaleCount = st.CountByYear
                        .ToDictionary(keySelector: k => (long)k.BirthYear, elementSelector: e => (int)e.CountBySex.Where(c => !(bool)c.ChildMale).Select(s => s.ChildCount).FirstOrDefault())
               ,
                AttMaleCount = st.CountBySex.Where(c => (bool)c.AttMale) //распределение по годам
                        .ToDictionary(keySelector: k => (long)k.BirthYear, elementSelector: e => (int)e.AttCount)
               ,
                AttFemaleCount = st.CountBySex.Where(c => !(bool)c.AttMale) //распределение по годам
                        .ToDictionary(keySelector: k => (long)k.BirthYear, elementSelector: e => (int)e.AttCount)
            })
        .ToList();

    //Делаем колонку всего под Key -1
    dataResult.Data.ToList().ForEach(d => d.MaleCount.Add(-1, d.MaleCount.Sum(m => m.Value)));
    dataResult.Data.ToList().ForEach(d => d.FemaleCount.Add(-1, d.FemaleCount.Sum(fm => fm.Value)));

    //переносим данные из Att(Female/Male)Count в (...)Count
    dataResult.Data.ToList().ForEach(d =>
    {
        d.AttMaleCount.ToList().ForEach(m => { d.MaleCount.Add(m.Key, m.Value); });
        d.AttFemaleCount.ToList().ForEach(m => { d.FemaleCount.Add(m.Key, m.Value); });
    });

    //Строка Итого
    dataResult.Data.Add(new FamilyTimeOfRestInfo()
    {
        HotelName = "Итого",
        TimeOfRestName = "-",
        MaleCount = dataResult.SubHeaders.Select(h => new KeyValuePair<long, int>(h.Key, dataResult.Data.Select(d => d.MaleCount.Where(mc => mc.Key == h.Key).Sum(mc => mc.Value)).Sum())).ToDictionary(keySelector: k => k.Key, elementSelector: e => e.Value),
        FemaleCount = dataResult.SubHeaders.Select(h => new KeyValuePair<long, int>(h.Key, dataResult.Data.Select(d => d.FemaleCount.Where(mc => mc.Key == h.Key).Sum(mc => mc.Value)).Sum())).ToDictionary(keySelector: k => k.Key, elementSelector: e => e.Value),
        AttendantsFemaleCount = dataResult.Data.Select(d => d.AttendantsFemaleCount).Sum(),
        AttendantsMaleCount = dataResult.Data.Select(d => d.AttendantsMaleCount).Sum(),
    });

    return dataResult;
}