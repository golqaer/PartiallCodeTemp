use RestChildAiso

/****** Скрипт для команды SelectTopNRows из среды SSMS  ******/
SELECT
		   hotel.Name as HotelName, tr.Name as TimeOfRestName
		, Count(ch.Id) as ChildCount
		, ch.Male as ChildMale
		, YEAR(ch.DateOfBirth) as BirthDate
	FROM [RestChildAiso].[dbo].Request as req
		join Child as ch on req.Id = ch.RequestId
		join TypeOfRest as tor on req.TypeOfRestId = tor.Id and tor.ParentId = 2 --совместный отдых
		join Tour as tour on tour.Id = req.TourId
		join Hotels as hotel on req.HotelsId = hotel.Id
		join TimeOfRest as tr on tr.Id = tour.TimeOfRestId
	where ch.DateOfBirth is not null
		and req.StatusId = 1075
		--and tr.YearOfRestId = 9 --2022
		and ch.IsDeleted = 0
	group by hotel.Name, tr.Name, ch.Male, YEAR(ch.DateOfBirth)
	order by HotelName, TimeOfRestName,BirthDate

select HotelName, TimeOfRestName, Count(AttId) as AttendantCount, AttMale
	from
		(select hotel.Name as HotelName, tr.Name as TimeOfRestName, app.Id as AttId, app.Male as AttMale
			from Request as req
				join Applicant as app on (req.ApplicantId = app.Id and app.IsAccomp = 1) and app.Male is not Null		 
				join TypeOfRest as tor on req.TypeOfRestId = tor.Id and tor.ParentId = 2 --совместный отдых
				join Tour as tour on tour.Id = req.TourId
				join Hotels as hotel on req.HotelsId = hotel.Id
				join TimeOfRest as tr on tr.Id = tour.TimeOfRestId
			where req.StatusId = 1075
				and tr.YearOfRestId = 9 --2022
				and app.IsDeleted = 0
		union
		select hotel.Name as HotelName, tr.Name as TimeOfRestName, app.Id as AttId, app.Male as AttMale
			from Request as req
				join Applicant as app on app.RequestId = req.Id and app.Male is not null 
				join TypeOfRest as tor on req.TypeOfRestId = tor.Id and tor.ParentId = 2 --совместный отдых
				join Tour as tour on tour.Id = req.TourId
				join Hotels as hotel on req.HotelsId = hotel.Id
				join TimeOfRest as tr on tr.Id = tour.TimeOfRestId
			where req.StatusId = 1075
				and tr.YearOfRestId = 9 --2022
				and app.IsDeleted = 0) as atts				
	--where HotelName like '%Альфа%'
	group by HotelName, TimeOfRestName, AttMale
	order by HotelName, TimeOfRestName