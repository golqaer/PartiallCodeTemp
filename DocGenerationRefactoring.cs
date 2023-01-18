/*
 Объемный рефакторинг, касающийся генерации документов pdf и doc, и то в полной мере не удалось провести из-за нехватки времени.
Задумка, была создать интерфейс и абстрактный класс, чтобы потом от них работать с реальными объектами по генерации doc и pdf. 
Объёмность задумки заключалась, в том что пришлось во всем проекте переписывать обращение к методам генерации доков, в иных местах
приходилось вникать очень глубоко в логику генерации основанную на автоматическом проставлении цепочки статусов заявлений
и логику бизнес процесса.
Сами методы создания документов пребывают в плачевном состоянии - много повторяющегося кода. Это проблема самой концепции создания документов.
 */

//интерфейс
using RestChild.DocumentGeneration.DocumentProc.Doc;
using RestChild.DocumentGeneration.DocumentProc.Pdf;
using RestChild.DocumentGeneration.DocumentProc;

namespace RestChild.DocumentGeneration.DocumentProc
{
	public interface IBaseDocumentProcessor
	{
		string Extension { get; }
		string MimeType { get; }

		/// <summary>
		/// Развилка для уведомления о регистрации
		/// </summary>
		IDocument NotificationRegistration(Request request, Account account);

		/// <summary>
		/// Сертификат (на отдых и пр)
		/// </summary>
		IDocument SaveCertificateToRequest(Request request, long? sendStatusId = null);

		/// <summary>
		/// Уведомление (базовое)
		/// </summary>
		IDocument NotificationRefuse(Request request, Account account);

		/// <summary>
		/// Уведомление о приостановлении рассмотрения заявления
		/// </summary>
		IDocument NotificationWaitApplicant(Request request, Account account);

		/// <summary>
		/// Уведомление о принятом решении
		/// </summary>
		IDocument NotificationAboutDecision(Request request, Account account);

		/// <summary>
		/// Уведомление о необходимости выбора организации отдыха и оздоровления
		/// </summary>
		IDocument NotificationOrgChoose(Request request, Account account);

		/// <summary>
		/// Уведомление о замене сопровождающего
		/// </summary>
		IDocument NotificationAttendantChange(Request request, Account account, long oldAttendantId = 0, long attendantId = 0, string changeDat = null);

		/// <summary>
		/// Уведомление об отказе от путёвки/сертификата
		/// </summary>
		IDocument NotificationDeclineAccepted(Request request, Account account);

		/// <summary>
		/// Уведомление об отказе в отказе от путёвки/сертификата
		/// </summary>
		IDocument NotificationDeclineNotAccepted(Request request, Account account);

		/// <summary>
		/// Уведомление о регистрации отказа в реестре отказов
		/// </summary>
		IDocument NotificationDeclineRegistryReg(Request request, Account account);
	}
}

//абстрактный класс
namespace RestChild.DocumentGeneration.DocumentProc
{
	static public class DocTypeEnum
	{
		/// <summary>
		/// Word - .docx
		/// </summary>
		public static readonly string Doc = "docx";

		/// <summary>
		/// Acrobat Reader - .pdf
		/// </summary>
		public static readonly string Pdf = "pdf";
	}

	public abstract class BaseDocumentProcessor : IBaseDocumentProcessor
	{
		protected string _extension;
		protected const string _SPACE = " ";

		protected const string _FEDERAL_LAW = "Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее – Порядок).";
		protected const string _FEDERAL_SHORT2020_LAW = "Порядок организации отдыха и оздоровления детей на основании путевок для отдыха и оздоровления при выездном отдыхе с полной оплатой стоимости путевки для отдыха и оздоровления за счет средств бюджета города Москвы с датами заезда, начиная с 1 апреля по 27 июля 2020 г., в другое время либо замены указанных путевок на сертификат на отдых и оздоровление";
		protected const string _FEDERAL_SHORT2021_LAW = "Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\"";
		protected const string _FEDERAL_LAW_REF = "пункты 8.9 и 9.17 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее – Порядок).";
		protected const string _PARTICIPATE_NOTIFICATION = "в случае неучастия заявителя, подавшего заявление о предоставлении услуг отдыха и оздоровления в отношении организации индивидуального выездного отдыха либо совместного выездного отдыха ребенка, находящегося в трудной жизненной ситуации и указанного в пункте 3.1.2 Порядка, лица из числа детей-сирот и детей, оставшихся без попечения родителей, во втором этапе заявочной кампании, услуга отдыха и оздоровления не считается оказанной.";
		protected const string _DECLINE_REASON_PARTICIPATE = "пункт 8.8. Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее – Порядок).";
		protected const string _PARTICIPATE = "В случае неучастия заявителя, подавшего заявление о предоставлении услуг отдыха и оздоровления в отношении организации индивидуального выездного отдыха либо совместного выездного отдыха ребенка, отнесенного к одной из категорий детей, находящихся в трудной жизненной ситуации и указанных в пунктах 3.1.3 и 3.1.4 Порядка, заявителя, подавшего заявление о предоставлении услуг отдыха и оздоровления в отношении организации индивидуального выездного отдыха ребенка, отнесенного к одной из категорий детей, находящихся в трудной жизненной ситуации и указанных в пунктах 3.1.5 - 3.1.13 Порядка, во втором этапе заявочной кампании, услуга отдыха и оздоровления считается предоставленной";


		public string Extension { get { return "." + _extension; } }
		public string MimeType { get; }

		public BaseDocumentProcessor(string extenstion)
		{
			this._extension = extenstion;
			this.MimeType = MimeTypeMap.GetMimeType(Extension);
		}

		/// <summary>
		/// Развилка для уведомления о регистрации
		/// </summary>
		public IDocument NotificationRegistration(Request request, Account account)
		{

			if (request.StatusId == (long)StatusEnum.RegistrationDecline)
			{
				return new CshedDocumentResult(NotificationRegistrationDecline(request, account));
			}
			if (request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest)
			{
				return new CshedDocumentResult(NotificationCompensationRegistration(request, account));
			}
			if (request.TypeOfRestId == (long)TypeOfRestEnum.Compensation)
			{
				if (request.Child?.FirstOrDefault().BenefitTypeId == (long)BenefitTypeEnum.CompensationOrphan)
				{
					return new CshedDocumentResult(NotificationCompensationRegistrationOrphans(request, account));
				}

				return new CshedDocumentResult(NotificationCompensationRegistrationForPoors(request, account));
			}

			return new CshedDocumentResult(NotificationBasicRegistration(request, account));
		}

		/// <summary>
		/// Сертификат (на отдых и пр)
		/// </summary>
		public IDocument SaveCertificateToRequest(Request request, long? sendStatusId = null)
		{

			if (request == null || request.IsDeleted || request.StatusId != (long)StatusEnum.CertificateIssued &&
					sendStatusId != (long)StatusEnum.CertificateIssued)
				return null;

			return new CshedDocumentResult(CertificateForRequestTemporaryFile(request, sendStatusId));
		}

		/// <summary>
		/// Уведомление (базовое)
		/// </summary>
		public IDocument NotificationRefuse(Request request, Account account)
		{
			var dr2 = ConfigurationManager.AppSettings["NotificationRefuseDeclineReasonWrongDocs"].LongParse() ?? 201904;
			var dr3 = ConfigurationManager.AppSettings["NotificationRefuseDeclineReasonQuota"].LongParse() ?? 201705;
			// отказ от вариантов
			var dr4 = ConfigurationManager.AppSettings["NotificationRefuseDeclineDiscardingOptions"].LongParse() ?? 201902;
			// неучастие в выборе альтернативного варианта (мини ЛОК 2020)
			var dr5 = ConfigurationManager.AppSettings["NotificationRefuseDeclineDiscardingChoose"].LongParse() ?? 201911;
			//неучастие заявителя во втором этапе заявочной кампании
			var dr6 = ConfigurationManager.AppSettings["NotParticipateInSecondStage"].LongParse() ?? 201911;


			if (request.StatusId == (long)StatusEnum.CancelByApplicant && !string.IsNullOrWhiteSpace(request.CertificateNumber))
			{
				var doc = NotificationDeadline(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return new CshedDocumentResult(doc);
			}
			if (request.StatusId == (long)StatusEnum.CancelByApplicant)
			{
				return new CshedDocumentResult((NotificationRefuse1090(request, account)));
			}
			if (request.StatusId == (long)StatusEnum.Reject && request.DeclineReasonId == dr2)
			{
				var doc = NotificationRefuse10802(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return new CshedDocumentResult(doc);
			}
			if (request.StatusId == (long)StatusEnum.Reject && request.DeclineReasonId == dr3)
			{
				var doc = NotificationRefuse10805(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return new CshedDocumentResult(doc);
			}
			if (request.StatusId == (long)StatusEnum.Reject && request.DeclineReasonId == dr6)
			{
				var doc = NotificationRefuse108013(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return new CshedDocumentResult(doc);
			}
			if (request.StatusId == (long)StatusEnum.Reject &&
				request.TypeOfRestId == (long)TypeOfRestEnum.Compensation)
			{
				var doc = NotificationRefuseCompensation(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return new CshedDocumentResult(doc);
			}
			if (request.StatusId == (long)StatusEnum.Reject &&
				request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest)
			{
				var doc = NotificationRefuseCompensationYouthRest(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return new CshedDocumentResult(doc);
			}
			if (request.StatusId == (long)StatusEnum.Reject)
			{
				var doc = NotificationRefuse1080(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return new CshedDocumentResult(doc);
			}

			var document = NotificationRefuseContent(request, account);
			document.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;

			return new CshedDocumentResult(document);
		}

		/// <summary>
		/// Уведомление о приостановлении рассмотрения заявления
		/// </summary>
		public IDocument NotificationWaitApplicant(Request request, Account account)
		{
			IEnumerable<Child> childs = request.Child; // unitOfWork.GetSet<Child>().Where(x => x.RequestId == requestId);
			IEnumerable<BenefitType> benefits = childs.Where(c => c.BenefitType.IsActive).Select(c => c.BenefitType); //= unitOfWork.GetAll<BenefitType>().Where(b => b.IsActive && childs.Any(c => c.BenefitTypeId == b.Id));

			return new CshedDocumentResult(NotificationWaitApplicant(request, account, benefits));
		}

		/// <summary>
		/// Уведомление о принятом решении
		/// </summary>
		public IDocument NotificationAboutDecision(Request request, Account account)
		{
			var money = new[]
				{
					(long?) TypeOfRestEnum.Money
					, (long?) TypeOfRestEnum.MoneyOn3To7
					, (long?) TypeOfRestEnum.MoneyOn7To15
					, (long?) TypeOfRestEnum.MoneyOn18
					, (long?) TypeOfRestEnum.MoneyOnInvalidOn4To17
					, (long?) TypeOfRestEnum.MoneyOnInvalidAndEscort4To17
					, (long?) TypeOfRestEnum.MoneyOnLimitationAndEscort4To17
                    //,(long?) TypeOfRestEnum.RestWithParentsOrphan
                    //,(long?) TypeOfRestEnum.YouthRestOrphanCamps
                    //,(long?) TypeOfRestEnum.ChildRestOrphanCamps
                };

			if (request.StatusId == (long)StatusEnum.CertificateIssued &&
				(request.TypeOfRestId == (long)TypeOfRestEnum.Compensation ||
				 request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest))
			{
				return new CshedDocumentResult(NotificationAboutCompensationIssued(request, account));
			}

			if (money.Contains(request.TypeOfRestId))
			{
				return new CshedDocumentResult(NotificationAboutCertificate(request, account));
			}


			return new CshedDocumentResult(NotificationAboutTour(request, account));
		}

		/// <summary>
		/// Уведомление о необходимости выбора организации отдыха и оздоровления
		/// </summary>
		public abstract IDocument NotificationOrgChoose(Request request, Account account);

		/// <summary>
		/// Уведомление о замене сопровождающего
		/// </summary>
		public IDocument NotificationAttendantChange(Request request, Account account, long oldAttendantId = 0, long attendantId = 0, string changeDate = null)
		{
			if (attendantId == 0)
				attendantId = request.Attendant.Where(att => att.IsAccomp & !att.IsDeleted).Max(att => att.Id);
			if (oldAttendantId == 0)
				oldAttendantId = request.Attendant.Where(att => att.IsDeleted & att.AttendantChangeId == attendantId).Max(att => (long?)att.Id)
						?? request.Applicant.Id;

			if (!string.IsNullOrEmpty(changeDate))
				changeDate = request.Attendant.Where(att => att.Id == attendantId).FirstOrDefault().AttendantChangeRequestDate.FormatEx("dd.MM.yyyy");

			return new CshedDocumentResult(InnerNotificationAttendantChange(request, account, oldAttendantId, attendantId, changeDate));
		}

		/// <summary>
		/// Уведомление об отказе от путёвки/сертификата
		/// </summary>
		public abstract IDocument NotificationDeclineAccepted(Request request, Account account);

		/// <summary>
		/// Уведомление об отказе в отказе от путёвки/сертификата
		/// </summary>
		public abstract IDocument NotificationDeclineNotAccepted(Request request, Account account);

		#region PrivateMethods

		/// <summary>
		/// Уведомление о смене сопровождающего
		/// </summary>
		abstract protected IDocument InnerNotificationAttendantChange(Request request, Account account, long oldAttendantId, long attendantId, string changeDate);

		/// <summary>
		///     Сертификаты, основная развилка
		/// </summary>
		abstract public IDocument CertificateForRequestTemporaryFile(Request request, long? sendStatusId = null);

		/// <summary>
		/// Уведомление о приостановлении рассмотрения (вызов)
		/// </summary>
		abstract protected IDocument NotificationWaitApplicant(Request request, Account account, IEnumerable<BenefitType> benefits);

		/// <summary>
		/// Уведомление о регистрации отказа в реестре отказов
		/// </summary>
		abstract public IDocument NotificationDeclineRegistryReg(Request request, Account account);

		#region NotificationAboutDecision

		/// <summary>
		/// Уведомления о предоставлении выплаты компенсации
		/// </summary>
		abstract protected IDocument NotificationAboutCompensationIssued(Request request, Account account);

		/// <summary>
		/// Уведомление о предоставлении сертификата (1075.2)
		/// </summary>
		abstract protected IDocument NotificationAboutCertificate(Request request, Account account);

		/// <summary>
		/// Уведомление о предоставлении путёвки (1075.1)
		/// </summary>
		abstract protected IDocument NotificationAboutTour(Request request, Account account, bool forPrint = false);
		#endregion

		#region NotificationRefuse

		/// <summary>
		/// Уведомление об отказе в предоставлении услуг отдыха и оздоровления в связи с неучастием заявителя
		/// Уведомление формируемое при не участии в выборе альтернативной путёвки 1075
		/// </summary>
		abstract protected IDocument NotificationRefuseContent(Request request, Account account);

		/// <summary>
		/// Уведомление об отказе в предоставлении услуг отдыха и оздоровления (базовый)
		/// </summary>
		abstract protected IDocument NotificationRefuse1080(Request request, Account account);

		/// <summary>
		/// Уведомление об отказе в предоставлении услуг отдыха и оздоровления (лица из числа детей-сирот и детей, оставшихся
		/// без попечения родителей)
		/// </summary>
		abstract protected IDocument NotificationRefuseCompensationYouthRest(Request request, Account account);

		/// <summary>
		/// Уведомление об отказе в компенсации за отдых и оздоровление (дети из малообеспеченных семей, дети-сироты и дети,
		/// оставшиеся без попечения родителей)
		/// </summary>
		abstract protected IDocument NotificationRefuseCompensation(Request request, Account account);

		/// <summary>
		/// Уведомление об отказе в предоставлении услуг отдыха и оздоровления в связи с неучастием заявителя во втором этапе заявочной кампании
		/// Уведомление формируемое при не участии в выборе альтернативной путёвки 1075
		/// </summary>
		abstract protected IDocument NotificationRefuse108013(Request request, Account account);

		/// <summary>
		/// Уведомление об отказе в предоставлении услуг отдыха (по квоте)
		/// </summary>
		abstract protected IDocument NotificationRefuse10805(Request request, Account account);

		/// <summary>
		/// уведомление об отказе в установленный срок
		/// </summary>
		abstract protected IDocument NotificationDeadline(Request request, Account account);

		/// <summary>
		/// Уведомление о прекращении рассмотрения поданного заявления (заявитель отказался сам)
		/// </summary>
		abstract protected IDocument NotificationRefuse1090(Request request, Account account);

		/// <summary>
		/// Уведомление об отказе в предоставлении услуг отдыха и оздоровления в связи с представлением документов, не
		/// соответствующих требованиям
		/// </summary>
		abstract protected IDocument NotificationRefuse10802(Request request, Account account);

		#endregion

		#region NotificationRegistration

		/// <summary>
		/// Базовое уведомление о регистрации
		/// </summary>
		/// <param name="request"></param>
		/// <param name="account"></param>
		/// <returns></returns>
		abstract protected IDocument NotificationBasicRegistration(Request request, Account account);

		/// <summary>
		///     Уведомление о отказе в регистрации заявления (1035)
		/// </summary>
		abstract protected IDocument NotificationRegistrationDecline(Request request, Account account);

		abstract protected IDocument NotificationCompensationRegistration(Request request, Account account);

		abstract protected IDocument NotificationCompensationRegistrationOrphans(Request request, Account account);

		abstract protected IDocument NotificationCompensationRegistrationForPoors(Request request, Account account);


		#endregion

		#endregion
	}
}

//PDF генератор, там очень много кода, нет времени выдергивать, все любезно разбито на регионы
namespace RestChild.DocumentGeneration.DocumentProc.Pdf
{
	public class PdfProc : BaseDocumentProcessor
	{
		#region ConstSettings
		#region fontsInit
		private static readonly BaseFont _BASE_FONT = BaseFont.CreateFont(
			Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\times.ttf", BaseFont.IDENTITY_H,
			BaseFont.EMBEDDED);

		private static readonly BaseFont _BASE_BOLD_FONT = BaseFont.CreateFont(
			Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\timesbd.ttf", BaseFont.IDENTITY_H,
			BaseFont.EMBEDDED);

		private static readonly BaseFont _BASE_ITALIC_FONT = BaseFont.CreateFont(
			Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\timesi.ttf", BaseFont.IDENTITY_H,
			BaseFont.EMBEDDED);

		/// <summary>
		/// Шрифт для ЭЦП
		/// </summary>
		private static readonly BaseFont _ECP_BASE_FONT = BaseFont.CreateFont(
			Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\arial.ttf", BaseFont.IDENTITY_H,
			BaseFont.EMBEDDED);

		/// <summary>
		/// Шрифт для ЭЦП
		/// </summary>
		private static readonly BaseFont _ECP_BASE_BOLD_FONT = BaseFont.CreateFont(
			Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\arialbd.ttf", BaseFont.IDENTITY_H,
			BaseFont.EMBEDDED);

		#region fontPathes
		private static readonly string _arialPath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\arial.ttf";
		private static readonly string _acromPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AcromLight.otf");
		private static readonly string _acromOtherPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin"), "AcromLight.otf");
		#endregion
		private static readonly BaseFont _customFont = File.Exists(_acromOtherPath) ? BaseFont.CreateFont(_acromOtherPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
																 : File.Exists(_acromPath)
																 ? BaseFont.CreateFont(_acromPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
																 : BaseFont.CreateFont(_arialPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

		private static readonly Font _FONT_8 = new Font(_customFont, 8);
		private static readonly Font _FONT_10 = new Font(_customFont, 10);
		private static readonly Font _FONT_12 = new Font(_customFont, 12);
		private static readonly Font _FONT_14 = new Font(_customFont, 14);
		#endregion

		/// <summary>
		/// Путь к иконке, которая отображается в ЭЦП
		/// </summary>
		private readonly string _ECP_ICON_PATH = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "eds-logo.png");

		private readonly Font _HEADER_FONT = new Font(_BASE_BOLD_FONT, 11);
		private readonly Font _TITLE_FONT = new Font(_BASE_BOLD_FONT, 14);
		private readonly Font _MAIN_TXT_FONT = new Font(_BASE_FONT, 11);
		private readonly Font _MAIN_TXT_ITALIC_FONT = new Font(_BASE_ITALIC_FONT, 11);
		private readonly Font _SMALL_TXT_FONT = new Font(_BASE_FONT, 8);
		private readonly Font _SMALL_TXT_FONT_UNDERLINED = new Font(_BASE_FONT, 11, Font.UNDERLINE);

		private readonly float _FIRST_LINE = 70;
		private readonly float _FIRST_LINE_SMALL = 30;
		#endregion

		public PdfProc() : base(DocTypeEnum.Pdf)
		{

		}


		#region Override

		public override IDocument NotificationDeclineNotAccepted(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };
			var isCert = request.TypeOfRestId == (long)TypeOfRestEnum.Money ||
								 request.TypeOfRest.ParentId == (long)TypeOfRestEnum.Money ||
								 //request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn18 || //TODO раскоментировать строку во время альтернативной компании(требуется уточнение)
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalid || почему-то были добавлены эти два типа отдыха, хотя они ни разу не сертификаты
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalidOrphanComplex ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidOn4To17 ||
								 request.TimeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidAndEscort4To17 ||
								 request.TimeOfRestId == (long)TypeOfRestEnum.MoneyOnLimitationAndEscort4To17 ||
								 request.TimeOfRestId == (long)TypeOfRestEnum.MoneyOnBenefits;

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);


					if (isCert)
					{
						PdfAddParagraph(document, "Уведомление о невозможности сформировать отказ от использования \nсертификата на отдых и оздоровление", 0,
						1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Сведения заявления об отказе от использования сертификата на отдых и оздоровление:", _HEADER_FONT));
					}
					else
					{
						PdfAddParagraph(document, "Уведомление о невозможности сформировать отказ от осуществления \nотдыха и оздоровления на основании предоставленной бесплатной \nпутевки для отдыха и оздоровления (в связи с подачей заявления после \nустановленного срока)", 0,
					   1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Сведения заявления об отказе от осуществления отдыха и оздоровления:", _HEADER_FONT));
					}

					SignCertBlock(document, request);

					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Номер сертификата: ", _HEADER_FONT),
						new Chunk(request.CertificateNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Номер путёвки: ", _HEADER_FONT),
						new Chunk(request.CertificateNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
						if (!request.Tour.IsNullOrEmpty())
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						   new Chunk("Организация отдыха и оздоровления: ", _HEADER_FONT),
						   new Chunk(request.Tour.Hotels.NameOrganization.FormatEx(), _MAIN_TXT_ITALIC_FONT));

						if (!request.TimeOfRest.IsNullOrEmpty())
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Время отдыха: ", _HEADER_FONT),
							new Chunk(request.Tour.DateIncome.FormatEx("dd.MM.yyyy") + " - " + request.Tour.DateOutcome.FormatEx("dd.MM.yyyy"), _MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					if (request.Child != null)
					{
						if (isCert)
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Данные лиц, указанных в сертификате: ", _HEADER_FONT));
						}
						else
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Данные лиц, указанных в путевке: ", _HEADER_FONT));
						}
						var attendants = new List<Applicant>();
						if (request.Applicant?.IsAccomp ?? false)
						{
							attendants.Add(request.Applicant);
						}
						if (request.Attendant?.Count > 0)
						{
							attendants.AddRange(request.Attendant.ToList());
						}
						foreach (var attendant in attendants.Where(att => !att.IsDeleted))
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk(
							$"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}, {attendant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
						}
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk(
							$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
						}
					}


					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("сформировать отказ от использованного сертификата на отдых и оздоровление не представляется возможным. Сертификат на отдых и оздоровление использован.", _MAIN_TXT_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Основание: ", _HEADER_FONT),
						new Chunk("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\".", _MAIN_TXT_FONT));

					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("заявление подано после установленного срока. Сформировать отказ от осуществления отдыха и оздоровления на основании предоставленной бесплатной путевки для отдыха и оздоровления не представляется возможным.", _MAIN_TXT_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Основание: ", _HEADER_FONT),
						new Chunk("пункт 10.1. Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56 - ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\". Заявление об отказе от предоставленной бесплатной путевки для отдыха и оздоровления в установленный порядком срок подается не позднее 35 рабочих дней до даты заезда в организацию отдыха и оздоровления.", _MAIN_TXT_FONT));

					}

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о непринятии отказа" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		public override IDocument NotificationDeclineAccepted(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };
			var isCert = request.TypeOfRestId == (long)TypeOfRestEnum.Money ||
								 request.TypeOfRest.ParentId == (long)TypeOfRestEnum.Money ||
								 //request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn18 || //TODO раскоментировать строку во время альтернативной компании(требуется уточнение)
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalid || почему-то были добавлены эти два типа отдыха, хотя они ни разу не сертификаты
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalidOrphanComplex ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidOn4To17 ||
								 request.TimeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidAndEscort4To17 ||
								 request.TimeOfRestId == (long)TypeOfRestEnum.MoneyOnLimitationAndEscort4To17 ||
								 request.TimeOfRestId == (long)TypeOfRestEnum.MoneyOnBenefits;

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);


					if (isCert)
					{
						PdfAddParagraph(document, "Уведомление о сформированном отказе от использования сертификата на отдых и оздоровление", 0,
						1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
						//PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						//new Chunk("Сведения заявления об отказе от использования сертификата на отдых и оздоровление:", _HEADER_FONT));
					}
					else
					{
						PdfAddParagraph(document, "Уведомление о сформированном отказе от осуществления отдыха\nи оздоровления на основании предоставленной бесплатной путевки\nдля отдыха и оздоровления", 0,
					   1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
						//PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						//new Chunk("Сведения заявления об отказе от осуществления отдыха и оздоровления:", _HEADER_FONT));
					}

					//SignCertBlock(document, request);

					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Номер сертификата: ", _HEADER_FONT),
						new Chunk(request.CertificateNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Номер путёвки: ", _HEADER_FONT),
						new Chunk(request.CertificateNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
						if (!request.Tour.IsNullOrEmpty())
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						   new Chunk("Организация отдыха и оздоровления: ", _HEADER_FONT),
						   new Chunk(request.Tour.Hotels.NameOrganization ?? request.Tour.Hotels.Name, _MAIN_TXT_ITALIC_FONT));

						if (!request.Tour.IsNullOrEmpty())
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Время отдыха: ", _HEADER_FONT),
							new Chunk(request.Tour.DateIncome.FormatEx("dd.MM.yyyy") + " - " + request.Tour.DateOutcome.FormatEx("dd.MM.yyyy"), _MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					if (request.Child != null)
					{
						if (isCert)
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Данные лиц, указанных в сертификате: ", _HEADER_FONT));
						}
						else
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Данные лиц, указанных в путевке: ", _HEADER_FONT));
						}
						var attendants = new List<Applicant>();
						if (request.Applicant?.IsAccomp ?? false)
						{
							attendants.Add(request.Applicant);
						}
						if (request.Attendant?.Count > 0)
						{
							attendants.AddRange(request.Attendant.ToList());
						}
						foreach (var attendant in attendants.Where(att => !att.IsDeleted))
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk(
							$"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}, {attendant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
						}
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk(
							$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
						}
					}


					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("Заявление об отказе от предоставленного сертификата на отдых и оздоровление удовлетворено. Сертификат на отдых и оздоровление отозван по инициативе заявителя.", _MAIN_TXT_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Основание: ", _HEADER_FONT),
						new Chunk("пункт 8(1).10 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утверждённого постановлением Правительства Москвы от 22 февраля 2017 г. № 56 - ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\".", _MAIN_TXT_FONT));

					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("Причина отказа от предоставленной путевки для отдыха и оздоровления признана уважительной. Отказ от осуществления отдыха и оздоровления на основании предоставленной бесплатной путевки для отдыха и оздоровления сформирован.", _MAIN_TXT_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Основание: ", _HEADER_FONT),
						new Chunk("пункты 10.1 и 10.2 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56 - ПП \"Об организации отдыха и оздоровления детей, находящихся \"в трудной жизненной ситуации\".", _MAIN_TXT_FONT));

					}

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в установленный срок" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationWaitApplicant(Request request, Account account, IEnumerable<BenefitType> benefits)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);
					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, "Уведомление о приостановлении рассмотрения заявления о предоставлении услуг отдыха и оздоровления", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));


					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей:",
								_HEADER_FONT),
							new Chunk("\n", _MAIN_TXT_ITALIC_FONT),
							new Chunk(
								$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps || request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей:",
								_HEADER_FONT),
							new Chunk("\n", _MAIN_TXT_ITALIC_FONT),
							new Chunk(
								$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk(
								"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет",
								_MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Этап рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("рассмотрение заявления о предоставлении услуг отдыха и оздоровления приостановлено.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Основание: ", _HEADER_FONT),
						new Chunk($"пункт 6.4 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\": ", _MAIN_TXT_FONT),
						new Chunk("\"Необходимость личной явки заявителя в ГАУК \"МОСГОРТУР\".", _HEADER_FONT));

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						new Chunk("Для подтверждения сведений, указанных в заявлении о предоставлении услуг отдыха и оздоровления (далее – заявление), ", _MAIN_TXT_FONT),
						new Chunk("в течение 10 рабочих дней со дня направления данного уведомления", _HEADER_FONT),
						new Chunk(" Вам необходимо явиться в офис ГАУК \"МОСГОРТУР\" по адресу: город Москва, Малый Харитоньевский переулок, дом 6, строение 3.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						new Chunk("Прием в офисе ГАУК \"МОСГОРТУР\" осуществляется исключительно по предварительной записи.", _HEADER_FONT),
						new Chunk(" Запись производится через официальный портал Мэра и Правительства Москвы mos.ru (далее – Портал) или при личном визите заявителя в офис ГАУК \"МОСГОРТУР\". Запись производится на свободную дату и время.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("При себе необходимо иметь:", _HEADER_FONT));

					//List<string> docs = new List<string>
					//{
					//    "документ, удостоверяющий личность заявителя;",
					//    "документы, подтверждающие, что заявитель является родителем ребёнка (свидетельство о рождении ребенка, в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения);",
					//    "документ, подтверждающий место жительства ребёнка, лица из числа детей-сирот и детей, оставшихся без попечения родителей, в городе Москве;",
					//    "документ, подтверждающий полномочия заявителя, сопровождающего лица (в случае организации совместного выездного отдыха) из числа законных представителей – родителей, опекунов, попечителей, приемных родителей, патронатных воспитателей ребёнка (договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие статус ребенка);",
					//   "документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
					//    "документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребёнка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);",
					//    "документ, подтверждающий отнесение ребёнка, к одной из категорий детей, находящихся в трудной жизненной ситуации и указанных в пунктах 3.1.3, 3.1.5 - 3.1.13 Порядка, лица из числа детей-сирот к категории лиц из числа детей-сирот и детей, оставшихся без попечения родителей (заключение медико-социальной экспертизы, заключение Центральной психолого-медико-педагогической комиссии города Москвы, справки уполномоченного учреждения социальной защиты населения города Москвы и/или федеральных органов);"
					//};

					//доработка для наркоманов
					List<string> innerListOrphans = new List<string>();
					List<string> innerListDisabled = new List<string>();
					List<string> innerListLowIncome = new List<string>();
					List<string> innerListSacrifice = new List<string>();
					List<string> innerListRefugee = new List<string>();
					List<string> innerListExtreme = new List<string>();
					List<string> innerListViolence = new List<string>();
					List<string> innerListInvalid = new List<string>();
					List<string> innerListTerror = new List<string>();
					List<string> innerListMilitary = new List<string>();
					List<string> innerListInvalidParents = new List<string>();
					List<string> innerListDeviant = new List<string>();
					List<string> innerListOrphansYouth = new List<string>();


					if (benefits.Count() > 0 || (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest))
					{


						foreach (BenefitType benefit in benefits)
						{
							if (benefit.ExnternalUid.Contains("52"))//дети-сироты и оставшиеся без попечения родителей...
							{
								innerListOrphans = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является законным представителем ребенка: договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие полномочия законного представителя ребенка;",
										"документ, удостоверяющий личность ребенка;",
										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий полномочия заявителя, сопровождающего лица (в случае организации совместного выездного отдыха) из числа законных представителей – опекунов, попечителей, приемных родителей, патронатных воспитателей ребенка(договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие статус ребенка);",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность).",

									};


							}
							if (benefit.ExnternalUid.Contains("24"))//дети-инвалиды, дети с ограниченными возможностями...
							{

								innerListDisabled = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем (законным представителем) ребенка: свидетельство о рождении ребенка*, договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие полномочия законного представителя ребенка;",
										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, удостоверяющий личность ребенка;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении (заключение медико-социальной экспертизы или заключение Центральной психолого-медико-педагогической комиссии города Москвы**);",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени;",
										"** является документом, подтверждающим категорию, только в случае указания на необходимость предоставления ребенку специальных условий обучения по адаптированной основной образовательной программе."
									};



							}
							if (benefit.ExnternalUid.Contains("48"))//дети из малообеспеченных семей
							{

								innerListLowIncome = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, удостоверяющий личность ребенка;",
										"документ, подтверждающий место жительства ребенка в городе Москве;",

										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}

							if (benefit.ExnternalUid.Contains("57,71,72"))
							{
								innerListSacrifice = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 7)
							{
								innerListRefugee = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 8)
							{
								innerListExtreme = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 9)
							{
								innerListViolence = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 10)
							{
								innerListInvalid = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 11)
							{
								innerListTerror = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.ExnternalUid.Contains("58,71,72"))
							{
								innerListMilitary = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.ExnternalUid.Contains("56"))
							{
								innerListInvalidParents = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 14)
							{
								innerListDeviant = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}

						}

						if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest)//лица из числа детей-сирот и детей, оставших без попечения...
						{


							innerListOrphansYouth = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документ, подтверждающий место жительства лица из числа детей-сирот и детей, оставшихся без попечения родителей, в городе Москве;",

										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"документ, подтверждающий отнесение лица из числа детей-сирот и детей, оставшихся без попечения родителей, к категории лиц из числа детей-сирот и детей, оставшихся без попечения родителей."

									};


						}



					}

					//if (benefits.Count() > 0 || (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest))
					//{
					//    docs.Clear();
					//    docs.Add("документ, удостоверяющий личность заявителя;");

					//    docs.Add("документ, подтверждающий место жительства ребенка в городе Москве;");
					//    docs.Add("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//    docs.Add("документы, подтверждающие, что заявитель является родителем (законным представителем) ребенка: свидетельство о рождении ребенка*, договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие полномочия законного представителя ребенка;");
					//    docs.Add("документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);");

					//    docs.Add("* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени");
					//    foreach (BenefitType benefit in benefits)
					//    {
					//        if (benefit.ExnternalUid.Contains("52"))//дети-сироты и оставшиеся без попечения родителей...
					//        {
					//            List<string> innerList = new List<string>
					//            {
					//                "документ, удостоверяющий личность ребенка;",
					//                "документ, подтверждающий полномочия заявителя, сопровождающего лица(в случае организации совместного выездного отдыха) из числа законных представителей – опекунов, попечителей, приемных родителей, патронатных воспитателей ребенка(договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие статус ребенка);",
					//                "документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);",

					//            };


					//            docs.InsertRange(4, innerList);
					//            docs.Remove("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//            docs.Remove("* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени");
					//        }
					//        if (benefit.ExnternalUid.Contains("24"))//дети-инвалиды, дети с ограниченными возможностями...
					//        {
					//            docs.Insert(1, "документ, удостоверяющий личность ребенка;");
					//            docs.Insert(3, "документ, подтверждающий отнесение ребенка к категории, указанной в заявлении (заключение медико-социальной экспертизы или заключение Центральной психолого-медико-педагогической комиссии города Москвы**);");
					//            docs.Insert(5, "документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);");
					//            docs.Remove("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//            docs.Add("** является документом, подтвержающим категорию, только в случае указания на необходимость предоставления ребенку специальных условий обучения по адаптированной основной образовательной программе");
					//        }
					//        if (benefit.ExnternalUid.Contains("48"))//дети из малообеспеченных семей
					//        {

					//            docs.Insert(1, "документ, удостоверяющий личность ребенка;");
					//            docs.Insert(5, "документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);");
					//            docs.Remove("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//        }

					//    }

					//    if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest)//лица из числа детей-сирот и детей, оставших без попечения...
					//    {

					//        docs.Insert(1, "документ, подтверждающий место жительства лица из числа детей-сирот и детей, оставшихся без попечения родителей, в городе Москве;");
					//        docs.Insert(2, "документ, подтверждающий отнесение лица из числа детей-сирот и детей, оставшихся без попечения родителей, к категории лиц из числа детей-сирот и детей, оставшихся без попечения родителей.");
					//        docs.Remove("документ, подтверждающий место жительства ребенка в городе Москве;");
					//        docs.Remove("документы, подтверждающие, что заявитель является родителем (законным представителем) ребенка: свидетельство о рождении ребенка*, договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие полномочия законного представителя ребенка;");
					//        docs.Remove("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//        docs.Remove("* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени");
					//    }
					//}

					//foreach (var docText in docs)
					//{
					//    PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
					//}

					////PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("документ, удостоверяющий личность заявителя;", _MAIN_TXT_ITALIC_FONT));
					////PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("документы, подтверждающие, что заявитель является родителем ребёнка (свидетельство о рождении ребёнка, в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения);", _MAIN_TXT_ITALIC_FONT));
					////PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("документ, подтверждающий место жительства ребёнка, лица из числа детей-сирот и детей, оставшихся без попечения родителей, в городе Москве;", _MAIN_TXT_ITALIC_FONT));
					////PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("документ, подтверждающий полномочия заявителя, сопровождающего лица (в случае организации совместного выездного отдыха) из числа законных представителей – родителей, опекунов, попечителей, приемных родителей, патронатных воспитателей ребёнка (договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие статус ребенка);", _MAIN_TXT_ITALIC_FONT));
					////PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);", _MAIN_TXT_ITALIC_FONT));
					////PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребёнка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);", _MAIN_TXT_ITALIC_FONT));
					////PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("документ, подтверждающий отнесение ребёнка, к одной из категорий детей, находящихся в трудной жизненной ситуации и указанных в пунктах 3.1.3, 3.1.5 - 3.1.13 Порядка, лица из числа детей-сирот к категории лиц из числа детей-сирот и детей, оставшихся без попечения родителей (заключение медико-социальной экспертизы, заключение Центральной психолого-медико-педагогической комиссии города Москвы, справки уполномоченного учреждения социальной защиты населения города Москвы и/или федеральных органов).", _MAIN_TXT_ITALIC_FONT));

					if (innerListOrphans.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети-сироты и дети, оставшиеся без попечения родителей, находящиеся под опекой, попечительством, в том числе в приёмной или патронатной семье", _HEADER_FONT));

						foreach (var docText in innerListOrphans)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));

						}

					}

					if (innerListDisabled.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети-инвалиды и дети с ограниченными возможностями здоровья", _HEADER_FONT));

						foreach (var docText in innerListDisabled)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListLowIncome.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети из малообеспеченных семей", _HEADER_FONT));

						foreach (var docText in innerListLowIncome)
						{


							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListSacrifice.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети - жертвы вооруженных и межнациональных конфликтов, экологических и техногенных катастроф, стихийных бедствий", _HEADER_FONT));

						foreach (var docText in innerListSacrifice)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListRefugee.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети из семей беженцев и вынужденных переселенцев", _HEADER_FONT));

						foreach (var docText in innerListRefugee)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListExtreme.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети, оказавшиеся в экстремальных условиях", _HEADER_FONT));

						foreach (var docText in innerListExtreme)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListViolence.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети - жертвы насилия", _HEADER_FONT));

						foreach (var docText in innerListViolence)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListInvalid.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети, жизнедеятельность которых объективно нарушена в результате сложившихся обстоятельств и которые не могут преодолеть данные обстоятельства самостоятельно или с помощью семьи", _HEADER_FONT));

						foreach (var docText in innerListInvalid)
						{


							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListTerror.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети, пострадавшие в результате террористических актов", _HEADER_FONT));

						foreach (var docText in innerListTerror)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListMilitary.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети из семей военнослужащих и приравненных к ним лиц, погибших или получивших увечья(ранения, травмы, контузии) при исполнении ими обязанностей военной службы или служебных обязанностей", _HEADER_FONT));

						foreach (var docText in innerListMilitary)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListInvalidParents.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети из семей, в которых оба или один родитель являются инвалидами", _HEADER_FONT));

						foreach (var docText in innerListInvalidParents)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListDeviant.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Дети с отклонениями в поведении", _HEADER_FONT));

						foreach (var docText in innerListDeviant)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}

					if (innerListOrphansYouth.Count > 0)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Лица из числа детей-сирот и детей, оставшихся без попечения родителей", _HEADER_FONT));

						foreach (var docText in innerListOrphansYouth)
						{

							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk(docText, _MAIN_TXT_ITALIC_FONT));
						}

					}


					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Неявка с соответствующими документами в течение срока приостановления рассмотрения заявления является основанием для отказа в предоставлении услуг отдыха и оздоровления.", _MAIN_TXT_FONT));


					//SignWorkerBlock(document, account);

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о приостановлении рассмотрения заявления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		#region NotificationRegistration

		override protected IDocument NotificationBasicRegistration(Request request, Account account)
		{
			var youth = request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps;

			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					if (youth)
					{
						//PdfAddParagraph(document, "", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
							new Chunk("Уведомление о регистрации заявления о предоставлении услуг отдыха", _TITLE_FONT),
							new Chunk("\n", _TITLE_FONT),
							new Chunk("и оздоровления лицу из числа детей-сирот и детей, оставшихся без попечения родителей", _TITLE_FONT));
					}
					else
					{
						//PdfAddParagraph(document, "Уведомление о регистрации заявления о предоставлении услуг отдыха и оздоровления", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
							new Chunk("Уведомление о регистрации заявления о предоставлении услуг отдыха", _TITLE_FONT),
							new Chunk("\n", _TITLE_FONT),
							new Chunk("и оздоровления", _TITLE_FONT));
					}

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(youth ? "бесплатная путевка для отдыха и оздоровления" : request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));

					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребёнка:",
								_HEADER_FONT),
							new Chunk("\n", _TITLE_FONT),
							new Chunk(
								$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
					}


					if (youth)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные лица из числа детей-сирот и детей, оставшихся без попечения родителей: ",
								_HEADER_FONT),
							new Chunk(
								$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
					}


					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						new Chunk("Ваше заявление о предоставлении услуг отдыха и оздоровления (далее – заявление) ", _MAIN_TXT_FONT),
						new Chunk("зарегистрировано", _HEADER_FONT),
						new Chunk(".", _MAIN_TXT_FONT));



					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						new Chunk("В случае если в срок, не превышающий 21 рабочий день с даты подачи, по Вашему заявлению ", _MAIN_TXT_FONT),
						new Chunk("не пришло уведомление о необходимости личной явки заявителя в офис ГАУК \"МОСГОРТУР\" или уведомление об отказе в предоставлении услуг отдыха и оздоровления", _HEADER_FONT),
						new Chunk($" по причинам, указанным в пунктах {(youth ? "9.1.2-9.1.4, 9.1.6" : "9.1.2-9.1.6, 9.1.8-9.1.11")} Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее – Порядок), ", _MAIN_TXT_FONT),
						new Chunk("Ваше заявление прошло все проверки", _HEADER_FONT),
						new Chunk(" и будет рассмотрено на заседании Комиссии ГАУК \"МОСГОРТУР\" по включению сведений в предварительный Реестр получателей услуг отдыха и оздоровления (далее – Комиссия), которое состоится 16 января 2023 г.", _MAIN_TXT_FONT)
					);


					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						new Chunk("По результатам заседания Комиссии не позднее ", _MAIN_TXT_FONT),
						new Chunk("6 февраля 2023 г.", _HEADER_FONT),
						new Chunk(" в Ваш \"личный кабинет\" на Портале Мэра и Правительства Москвы будет направлено", _MAIN_TXT_FONT),
						new Chunk(youth ?
							" уведомление о необходимости выбора организации отдыха и оздоровления (допуск на второй этап заявочной кампании)." :
							" одно из уведомлений: о необходимости выбора организации отдыха и оздоровления (допуск на второй этап заявочной кампании); о предоставлении сертификата на отдых и оздоровление; об отказе в предоставлении услуг отдыха и оздоровления по причине, указанной в пункте 9.1.1 Порядка.", _MAIN_TXT_FONT));

					if (!youth)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
							new Chunk("При рассмотрении заявлений Комиссия руководствуется пунктом 3.9 Порядка, согласно которому в первую очередь рассматриваются заявления, в которых указан хотя бы один ребёнок, которому в период с 2020 по 2022 год не предоставлялись бесплатные путевки для отдыха и оздоровления (далее – бесплатные путевки) либо сертификаты на отдых и оздоровление (далее – сертификаты), а также не выплачивалась компенсация за самостоятельно приобретенную путевку (далее – компенсация).", _MAIN_TXT_FONT));

						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
							new Chunk("Во вторую очередь рассматриваются заявления, где указаны дети, которым в любые два года из последних трех лет (с 2020 по 2022 год) не предоставлялись бесплатные путевки либо сертификаты, а также не выплачивалась компенсация.", _MAIN_TXT_FONT));

						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
							new Chunk("В третью очередь рассматриваются заявления, где указаны дети, которым в любые два года из последних трех лет (с 2020 по 2022 год) предоставлялись бесплатные путевки либо сертификаты, а также выплачивалась компенсация.", _MAIN_TXT_FONT));

						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
							new Chunk("Исключение составляют дети-сироты и дети, оставшиеся без попечения родителей, находящиеся под опекой, попечительством, в том числе в приемной или патронатной семье, и лица из числа детей-сирот и детей, оставшихся без попечения родителей, которым федеральным законодательством гарантировано ежегодное предоставление услуг отдыха и оздоровления.", _MAIN_TXT_FONT));

						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
							new Chunk("По результатам второго этапа заявочной кампании, проводимого в период с 7 по 21 февраля 2023 г., уведомление о результате рассмотрения заявления (о предоставлении бесплатной путевки для отдыха и оздоровления; об отказе в предоставлении услуг отдыха и оздоровления в связи с неучастием заявителя во втором этапе заявочной кампании; о неучастии заявителя в выборе конкретной организации отдыха и оздоровления; об отказе в предоставлении услуг отдыха и оздоровления) направляется в \"личный кабинет\" на Портале Мэра и Правительства Москвы ", _MAIN_TXT_FONT),
							new Chunk("не позднее 22 февраля 2023 г.", _HEADER_FONT));

						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
							new Chunk("Обращаем внимание, что Вы согласились с необходимостью соблюдения: Правил отдыха и оздоровления, Соглашения об осуществлении обязанностей по сопровождению детей во время отдыха и оздоровления родителем или иным законным представителем либо доверенным лицом для сопровождения во время отдыха и оздоровления (для совместного выездного отдыха); а также проинформированы: о перечне противопоказаний к отдыху и оздоровлению, утвержденном Министерством здравоохранения Российской Федерации (для индивидуального выездного отдыха), об отсутствии возможности замены сертификата на отдых и оздоровление на бесплатную путевку для отдыха и оздоровления (в случае выбора сертификата на отдых и оздоровление), об отсутствии возможности замены бесплатной путевки для отдыха и оздоровления на сертификат на отдых и оздоровление (в случае выбора бесплатной путевки для отдыха и оздоровления), об отсутствии возможности выбора сертификата на отдых и оздоровление в случае отказа на втором этапе заявочной кампании от всех предложенных организаций отдыха и оздоровления, о перечислении денежных средств организации, осуществляющей туроператорскую деятельность и (или) оказывающей услуги отдыха и оздоровления, при приобретении у такой организации услуг отдыха и оздоровления или туристского продукта для ребёнка, либо ребенка и лица, его сопровождающего с использованием сертификата на отдых и оздоровление (в случае выбора сертификата на отдых и оздоровление). ", _MAIN_TXT_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
							new Chunk("По результатам второго этапа заявочной кампании, проводимого в период с 7 по 21 февраля 2023 г., уведомление о результате рассмотрения заявления (о предоставлении бесплатной путевки для отдыха и оздоровления, об отказе в предоставлении услуг отдыха и оздоровления в связи с неучастием заявителя во втором этапе заявочной кампании; об отказе в предоставлении услуг отдыха и оздоровления) будет направлено в Ваш \"личный кабинет\" на Портале Мэра и Правительства Москвы ", _MAIN_TXT_FONT),
							new Chunk("не позднее 22 февраля 2023 г.", _HEADER_FONT));

						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
							new Chunk("Обращаем внимание, что Вы согласились с необходимостью соблюдения Правил отдыха и оздоровления.", _MAIN_TXT_FONT));
					}

					//SignWorkerBlock(document, account);
					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о регистрации" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRegistrationDecline(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);
					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document,
						"Уведомление об отказе в регистрации заявления о предоставлении услуг отдыха и оздоровления",
						0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Дата обращения: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные обратившегося лица: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Цель обращения: ", _HEADER_FONT),
						new Chunk("получение услуг отдыха и оздоровления", _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));

					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей:\n",
								_HEADER_FONT),
							new Chunk(
								$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps || request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей:\n",
								_HEADER_FONT),
							new Chunk(
								$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Основание отказа в регистрации заявления: ", _HEADER_FONT),
						new Chunk("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее — Порядок).", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Причина отказа в регистрации заявления: ", _HEADER_FONT),
						//new Chunk("пункт 5.11.1 Порядка: \"Наличие в отношении одного и того же ребёнка, лица из числа детей - сирот и детей, оставшихся без попечения родителей, другого заявления о предоставлении услуг отдыха и оздоровления в текущем календарном году\".", _MAIN_TXT_FONT));
						new Chunk(request.DeclineReason.Name, _MAIN_TXT_FONT));
					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в регистрации заявления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};

			}
		}

		protected override IDocument NotificationCompensationRegistration(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					//PdfAddParagraph(document, "", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
					PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
						new Chunk("Уведомление о регистрации заявления на выплату компенсации", _TITLE_FONT),
						new Chunk("\n", _TITLE_FONT),
						new Chunk("за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления", _TITLE_FONT));



					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));

					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные лица из числа детей-сирот, и детей оставшихся без попечения родителей: ",
								_HEADER_FONT),
							new Chunk("\n", _TITLE_FONT),
							new Chunk(
								$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk($"{request.TypeOfRest.BenefitTypes.FirstOrDefault().Name}", _MAIN_TXT_ITALIC_FONT));
					}

					if (request.Child == null || !request.Child.Any())
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							 new Chunk(
								 "Данные лица из числа детей-сирот, и детей оставшихся без попечения родителей: ",
								 _HEADER_FONT),
							 new Chunk(
								 $"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								 _MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk($"{request.TypeOfRest.BenefitTypes.FirstOrDefault().Name}", _MAIN_TXT_ITALIC_FONT));
					}


					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Перечень предоставленных заявителем документов: ", _HEADER_FONT));

					var docs = new List<string>
					{
						"документ, удостоверяющий личность обратившегося лица",
						"страховой номер обязательного пенсионного страхования (СНИЛС) лица из числа детей-сирот и детей, оставшихся без попечения родителей",
						"доверенность, подтверждающая полномочия представителя, уполномоченного лицом из числа детей-сирот и детей, оставшихся без попечения родителей, на подачу заявления на выплату компенсации за самостоятельно приобретенную путевку (в случае подачи заявления на выплату компенсации за самостоятельно приобретенную путевку таким представителем)",
						"страховой номер обязательного пенсионного страхования (СНИЛС) представителя, уполномоченного лицом из числа детей-сирот и детей, оставшихся без попечения родителей, на подачу заявления на выплату компенсации за самостоятельно приобретенную путевку (в случае подачи заявления на выплату компенсации за самостоятельно приобретенную путевку таким представителем)",
						"документы, подтверждающие отдых и оздоровление, и оплату услуг отдыха и оздоровления лицом из числа детей-сирот и детей, оставшихся без попечения родителей",
						"сведения о кредитной организации и открытом лица из числа детей-сирот и детей, оставшихся без попечения родителей для выплаты компенсации за самостоятельно приобретенную лицом их числа детей-сирот и детей, оставшихся без попечения родителей путевку для отдыха и оздоровления"
					};

					AddTableDocsList(document, docs);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						new Chunk("Заявление", _HEADER_FONT),
						new Chunk(" на выплату компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления ", _MAIN_TXT_FONT),
						new Chunk("является неотъемлемой частью данного уведомления", _HEADER_FONT),
						new Chunk(".", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
					  new Chunk("О результате рассмотрения Вы будете уведомлены способом и формой информирования, указанной в заявлении.", _MAIN_TXT_FONT));

					//SignWorkerBlock(document, account);
					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о регистрации" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationCompensationRegistrationForPoors(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					//PdfAddParagraph(document, "", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
					PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
						new Chunk("Уведомление о регистрации заявления на выплату компенсации", _TITLE_FONT),
						new Chunk("\n", _TITLE_FONT),
						new Chunk("за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления", _TITLE_FONT),
						new Chunk("\n", _TITLE_FONT),
						new Chunk("(дети из малообеспеченных семей)", _TITLE_FONT));


					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));

					if (request.Child != null)
					{
						foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
								new Chunk(
									"Данные ребёнка: ",
									_HEADER_FONT),
								new Chunk(
									$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
									_MAIN_TXT_ITALIC_FONT));
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
								new Chunk("Льготная категория ребёнка: ", _HEADER_FONT),
								new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
						}
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Перечень предоставленных заявителем документов: ", _HEADER_FONT));

					var docs = new List<string>
					{
						"заявление на выплату компенсации за самостоятельно приобретенную путевку",
						"документ, удостоверяющий личность обратившегося лица",
						"документ, подтверждающий полномочия законного представителя – опекуна, попечителя, приемного родителя, патронатного воспитателя (в случае подачи заявления на выплату компенсации за приобретенную путевку обратившимся лицом из числа законных представителей – опекунов, попечителей, приемных родителей, патронатных воспитателей ребёнка)",
						"документ, удостоверяющий личность ребёнка:\nдля ребёнка в возрасте до 14 лет – свидетельство о рождении ребенка или документ, подтверждающий факт рождения и регистрации ребенка, выданный в установленном порядке (в случае рождения ребенка на территории иностранного государства)\nдля ребёнка, достигшего возраста 14 лет, – паспорт гражданина Российской Федерации, паспорт гражданина иностранного государства (в случае наличия гражданства иностранного государства)",
						"страховой номер обязательного пенсионного страхования (СНИЛС) ребёнка",
						"страховой номер обязательного пенсионного страхования (СНИЛС) родителя, иного законного представителя",
						"документы, подтверждающие факт родства родителя с ребёнком (в случае если иные представленные документы не содержат необходимой информации)",
						"документ, содержащий сведения о месте жительства ребёнка в городе Москве (в случае если в документе, удостоверяющем личность ребенка, отсутствуют сведения о его месте жительства в городе Москве)",
						"документы, подтверждающие отдых и оздоровление ребёнка и оплату родителем услуг отдыха и оздоровления",
						"сведения о кредитной организации и открытом счете родителя, иного законного представителя в кредитной организации для выплаты компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления",
						"доверенность, подтверждающая полномочия доверенного лица родителя (в случае подачи заявления на получение компенсации за приобретенную путевку доверенным лицом родителя, иного законного представителя ребёнка)"
					};

					AddTableDocsList(document, docs);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						new Chunk("Заявление", _HEADER_FONT),
						new Chunk(" на выплату компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления (дети из малообеспеченных семей) ", _MAIN_TXT_FONT),
						new Chunk("является неотъемлемой частью данного уведомления", _HEADER_FONT),
						new Chunk(".", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
					   new Chunk("О результате рассмотрения Вы будете уведомлены способом и формой информирования, указанной в заявлении.", _MAIN_TXT_FONT));


					//SignWorkerBlock(document, account);
					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о регистрации" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationCompensationRegistrationOrphans(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					//PdfAddParagraph(document, "", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
					PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
						new Chunk("Уведомление о регистрации заявления на выплату компенсации", _TITLE_FONT),
						new Chunk("\n", _TITLE_FONT),
						new Chunk("за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления", _TITLE_FONT),
						new Chunk("\n", _TITLE_FONT),
						new Chunk("(дети-сироты и дети, оставшиеся без попечения родителей, находящиеся под опекой, попечительством, в том числе", _TITLE_FONT),
						new Chunk("\n", _TITLE_FONT),
						new Chunk("в приемной или патронатной семье)", _TITLE_FONT));


					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));

					if (request.Child != null)
					{
						foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
								new Chunk(
									"Данные ребёнка: ",
									_HEADER_FONT),
								new Chunk(
									$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
									_MAIN_TXT_ITALIC_FONT));
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
								new Chunk("Льготная категория: ", _HEADER_FONT),
								new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
						}
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Перечень предоставленных заявителем документов: ", _HEADER_FONT));

					var docs = new List<string>
					{
						"заявление на выплату компенсации за самостоятельно приобретенную путевку",
						"документ, удостоверяющий личность обратившегося лица",
						"документ, подтверждающий полномочия законного представителя – опекуна, попечителя, приемного родителя, патронатного воспитателя",
						"документ, удостоверяющий личность ребёнка-сироты, ребенка, оставшегося без попечения родителей:\nдля ребёнка в возрасте до 14 лет – свидетельство о рождении ребенка или документ, подтверждающий факт рождения и регистрации ребенка, выданный в установленном порядке (в случае рождения ребенка на территории иностранного государства)\nдля ребёнка, достигшего возраста 14 лет, – паспорт гражданина Российской Федерации, паспорт гражданина иностранного государства (в случае наличия гражданства иностранного государства)",
						"страховой номер обязательного пенсионного страхования (СНИЛС) ребёнка-сироты, ребенка, оставшегося без попечения родителей",
						"страховой номер обязательного пенсионного страхования (СНИЛС) законного представителя ребёнка-сироты, ребенка, оставшегося без попечения родителей",
						"страховой номер обязательного пенсионного страхования (СНИЛС) сопровождающего лица (в случае сопровождения ребёнка-сироты, ребенка, оставшегося без попечения родителей во время отдыха и оздоровления)",
						"документы, подтверждающие отдых и оздоровление ребёнка-сироты, ребенка, оставшегося без попечения родителей, и оплату законным представителем услуг отдыха и оздоровления",
						"сведения о кредитной организации и открытом счете законного представителя в кредитной организации для выплаты компенсации за самостоятельно приобретенную законными представителями путевку для отдыха и оздоровления",
						"доверенность, подтверждающая полномочия доверенного лица законного представителя (в случае подачи заявления на получение компенсации за приобретенную путевку доверенным лицом законного представителя ребёнка)"
					};

					AddTableDocsList(document, docs);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						new Chunk("Заявление", _HEADER_FONT),
						new Chunk(" на выплату компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления (дети - сироты и дети, оставшиеся без попечения родителей, находящиеся под опекой, попечительством, в том числе в приемной или патронатной семье) ", _MAIN_TXT_FONT),
						new Chunk("является неотъемлемой частью данного уведомления", _HEADER_FONT),
						new Chunk(".", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
					   new Chunk("О результате рассмотрения Вы будете уведомлены способом и формой информирования, указанной в заявлении.", _MAIN_TXT_FONT));


					//SignWorkerBlock(document, account);
					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о регистрации" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		#endregion

		public override IDocument CertificateForRequestTemporaryFile(Request request, long? sendStatusId = null)

		{
			if (request.TypeOfRestId != (long)TypeOfRestEnum.YouthRestOrphanCamps &&
				request.TypeOfRestId != (long)TypeOfRestEnum.MoneyOn18 ||
				request.StatusId != (long)StatusEnum.CertificateIssued &&
				sendStatusId != (long)StatusEnum.CertificateIssued)
			{
				if (request.Child == null || request.Child.All(c => c.IsDeleted) ||
					request.StatusId != (long)StatusEnum.CertificateIssued &&
					sendStatusId != (long)StatusEnum.CertificateIssued)
				{
					return null;
				}
			}

			byte[] result;
			string fileName;

			using (var memoryStream = new MemoryStream())
			{
				var multiTypeRequest = request.Child.Count(c => !c.IsDeleted) > 1;

				if (request.RequestOnMoney)
				{
					if (!multiTypeRequest)
					{
						fileName = CertificateOnMoney(request, memoryStream);
					}
					else
					{
						fileName = CertificateOnMoneyMulti(request, memoryStream);
					}
				}
				else
				{
					if (!multiTypeRequest && request.CountAttendants <= 1)
					{
						fileName = CertificateOnRestSingle(request, memoryStream);
					}
					else
					{
						fileName = CertificateOnRestMulti(request, memoryStream);
					}
				}

				result = memoryStream.ToArray();
			}

			return new CshedDocumentResult
			{
				FileName = fileName,
				FileBody = result,
				MimeTypeShort = Extension,
				MimeType = MimeType
			};
		}

		#region NotificationRefuse

		protected override IDocument NotificationDeadline(Request request, Account account)

		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };
			var isCert = request.TypeOfRestId == (long)TypeOfRestEnum.Money ||
								 request.TypeOfRest.ParentId == (long)TypeOfRestEnum.Money ||
								 //request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn18 || //TODO раскоментировать строку во время альтернативной компании(требуется уточнение)
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalid || почему-то были добавлены эти два типа отдыха, хотя они ни разу не сертификаты
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalidOrphanComplex ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidOn4To17 ||
								 request.TimeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidAndEscort4To17 ||
								 request.TimeOfRestId == (long)TypeOfRestEnum.MoneyOnLimitationAndEscort4To17 ||
								 request.TimeOfRestId == (long)TypeOfRestEnum.MoneyOnBenefits;

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);


					if (isCert)
					{
						PdfAddParagraph(document, "Уведомление о сформированном отказе от использования сертификата на отдых и оздоровление", 0,
						1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
						//PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						//new Chunk("Сведения заявления об отказе от использования сертификата на отдых и оздоровление:", _HEADER_FONT));
					}
					else
					{
						PdfAddParagraph(document, "Уведомление о сформированном отказе от осуществления отдыха\nи оздоровления на основании предоставленной бесплатной путевки\nдля отдыха и оздоровления", 0,
					   1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
						//PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						//new Chunk("Сведения заявления об отказе от осуществления отдыха и оздоровления:", _HEADER_FONT));
					}

					//SignCertBlock(document, request);

					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Номер сертификата: ", _HEADER_FONT),
						new Chunk(request.CertificateNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Номер путёвки: ", _HEADER_FONT),
						new Chunk(request.CertificateNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
						if (!request.Tour.IsNullOrEmpty())
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						   new Chunk("Организация отдыха и оздоровления: ", _HEADER_FONT),
						   new Chunk(request.Tour.Hotels.NameOrganization ?? request.Tour.Hotels.Name, _MAIN_TXT_ITALIC_FONT));

						if (!request.Tour.IsNullOrEmpty())
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Время отдыха: ", _HEADER_FONT),
							new Chunk(request.Tour.DateIncome.FormatEx("dd.MM.yyyy") + " - " + request.Tour.DateOutcome.FormatEx("dd.MM.yyyy"), _MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					if (request.Child != null)
					{
						if (isCert)
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Данные лиц, указанных в сертификате: ", _HEADER_FONT));
						}
						else
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Данные лиц, указанных в путевке: ", _HEADER_FONT));
						}
						var attendants = new List<Applicant>();
						if (request.Applicant?.IsAccomp ?? false)
						{
							attendants.Add(request.Applicant);
						}
						if (request.Attendant?.Count > 0)
						{
							attendants.AddRange(request.Attendant.ToList());
						}
						foreach (var attendant in attendants.Where(att => !att.IsDeleted))
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk(
							$"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}, {attendant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
						}
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk(
							$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
						}
					}


					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("Заявление об отказе от предоставленного сертификата на отдых и оздоровление удовлетворено. Сертификат на отдых и оздоровление отозван по инициативе заявителя.", _MAIN_TXT_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Основание: ", _HEADER_FONT),
						new Chunk("пункт 8(1).10 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утверждённого постановлением Правительства Москвы от 22 февраля 2017 г. № 56 - ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\".", _MAIN_TXT_FONT));

					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("Причина отказа от предоставленной путевки для отдыха и оздоровления признана уважительной. Отказ от осуществления отдыха и оздоровления на основании предоставленной бесплатной путевки для отдыха и оздоровления сформирован.", _MAIN_TXT_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Основание: ", _HEADER_FONT),
						new Chunk("пункты 10.1 и 10.2 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56 - ПП \"Об организации отдыха и оздоровления детей, находящихся \"в трудной жизненной ситуации\".", _MAIN_TXT_FONT));

					}

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в установленный срок" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse1090(Request request, Account account)

		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);
					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, "Уведомление о прекращении рассмотрения поданного заявления\nо предоставлении услуг отдыха и оздоровления", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Причина обращения: ", _HEADER_FONT),
						new Chunk("прекращение рассмотрения поданного заявления о предоставлении услуг отдыха и оздоровления по инициативе заявителя.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Основание обращения: ", _HEADER_FONT),
						new Chunk($"пункт 5.13 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\": \"Заявитель вправе отозвать заявление о предоставлении услуг отдыха и оздоровления в срок не позднее дня окончания стадии заявочной кампании по приему заявлений о предоставлении услуг отдыха и оздоровления (до 10 декабря {DateTime.Now.Year} г. включительно)\".", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Информация о заявлении:", _HEADER_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE, new Chunk("дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE, new Chunk("номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("отозвано по инициативе заявителя в установленный срок.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Дата прекращения рассмотрения заявления: ", _HEADER_FONT),
						new Chunk(request.DateChangeStatus.FormatEx(string.Empty), _MAIN_TXT_FONT));

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о прекращении рассмотрения поданного заявления о предоставлении услуг отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse10802(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, "Уведомление об отказе в предоставлении услуг отдыха и оздоровления\nв связи с представлением документов, не соответствующих требованиям", 0,
						1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));
					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
								new Chunk(
									"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ",
									_HEADER_FONT),
								new Chunk(
									$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
									_MAIN_TXT_ITALIC_FONT));
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
								new Chunk("Льготная категория: ", _HEADER_FONT),
								new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
						}
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей:",
								_HEADER_FONT),
							new Chunk("\n", _MAIN_TXT_FONT),
							new Chunk(
								$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk(
								"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет",
								_MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("отказ в предоставлении услуг отдыха и оздоровления.", _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Основание отказа: ", _HEADER_FONT),
						new Chunk(_FEDERAL_LAW, _MAIN_TXT_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Причина отказа: ", _HEADER_FONT),
						new Chunk(request.DeclineReason?.Name, _MAIN_TXT_FONT));

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, "Уважаемый заявитель!", 0, 1, Element.ALIGN_CENTER, 0, _HEADER_FONT);
					PdfAddParagraph(document, "Просим обратить внимание на следующую информацию.", 0, 1, Element.ALIGN_CENTER, 0, _HEADER_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, "В случае принятия решения о повторной подаче заявления, рекомендуем убедиться в корректности внесенных в заявление сведений.", 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, _MAIN_TXT_ITALIC_FONT);
					PdfAddParagraph(document, "Обращаем внимание, что сведения, указанные в заявлении, должны полностью соответствовать сведениям, указанным в СНИЛС и документе, удостоверяющем личность заявителя, ребёнка/детей (в том числе следует обратить внимание на соответствие написания букв \"е-ё\").", 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, _MAIN_TXT_ITALIC_FONT);
					PdfAddParagraph(document, "В случае если сведения о Вас и ребёнке/детях внесены в \"личный кабинет\" Портала Мэра и Правительства Москвы (при подаче заявления данные отразились в полях автоматически), рекомендуем Вам проверить корректность внесенных сведений, а именно: сверить данные, указанные в \"личном кабинете\" Портала Мэра и Правительства Москвы и данные, указанные в СНИЛС и документе, удостоверяющем личность: фамилия, имя, отчество, пол, дата рождения.", 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, _MAIN_TXT_ITALIC_FONT);
					PdfAddParagraph(document, "В случае если Вами обнаружена ошибка в \"личном кабинете\" Портала Мэра и Правительства Москвы, ее необходимо исправить.", 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, _MAIN_TXT_ITALIC_FONT);
					PdfAddParagraph(document, "Дополнительно сообщаем, что в случае обнаружения расхождений сведений в СНИЛС и документе, удостоверяющем личность, Вам необходимо актуализировать документы в соответствующих органах, приведя их к единообразию.", 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, _MAIN_TXT_ITALIC_FONT);

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в предоставлении услуг отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse10805(Request request, Account account)
		{
			IUnitOfWork unitOfWork = new UnitOfWork();
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			var lokYear = request.YearOfRest?.Year ?? 2021;
			var yearIds = unitOfWork.GetSet<YearOfRest>().Where(ss => ss.Year < lokYear).OrderByDescending(ss => ss.Year)
														 .Take(3).Select(ss => ss.Id).ToList();
			IEnumerable<int> years = unitOfWork.GetSet<YearOfRest>().Where(x => yearIds.Contains(x.Id)).Select(x => x.Year).OrderBy(x => x).ToList();
			var listTravelersRequest = unitOfWork.GetSet<ListTravelersRequest>();

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, "Уведомление о результатах рассмотрения заявления о предоставлении услуг отдыха и оздоровления", 0,
						1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));


					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk("бесплатная путевка для отдыха и оздоровления/сертификат на отдых и оздоровление", _MAIN_TXT_ITALIC_FONT));

					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
							new Chunk(
								"Данные ребёнка:",
								_HEADER_FONT),
							new Chunk("\n", _MAIN_TXT_FONT),
							new Chunk(
								$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
							new Chunk(
								"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ",
								_HEADER_FONT),
							new Chunk(
								$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk(
								"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет",
								_MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk($"учитывая получение услуг отдыха и оздоровления в предыдущие годы и исходя из даты и времени подачи заявления, оказание услуг отдыха и оздоровления в 2023 году не представляется возможным.", _MAIN_TXT_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Основание: ", _HEADER_FONT),
						new Chunk("пункты 3.9. и 9.1.1. Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\".", _MAIN_TXT_FONT));


					var details = listTravelersRequest?.SelectMany(d => d.Details)
							   .Where(ss => ss.Detail != "[]")
							   .Select(ss => ss.Detail)
							   .SelectMany(JsonConvert.DeserializeObject<DetailInfo[]>).ToList();

					var _FIRST_LINE = true;

					IEnumerable<Request> requests = new List<Request>();
					foreach (var child in request.Child.Where(c => !c.IsDeleted).ToList())
					{
						// var requestIds = details.Where(ss => ss.ChildId == child.Id).Select(ss => ss.Id).Distinct().ToArray();
						//сравнение по снилсам потому что структура БД наркоманская
						var sameChilds = unitOfWork.GetSet<Child>().Where(ch => ch.Snils == child.Snils && ch.IsDeleted != true).ToList();

						var requestIds = sameChilds?.Select(ss => ss.RequestId).Distinct().ToList();

						requestIds.Remove(request.Id);
						if (requestIds.Count > 0)
						{

							requestIds.Remove(request.Id);


							requests = unitOfWork.GetSet<Request>().Where(re => requestIds.Any(req => req == re.Id)).ToList();

							var initialRequests = requests.Where(re => re.StatusId == (long)StatusEnum.Reject).ToList();

							var descendantRequests = unitOfWork.GetSet<Request>().Where(ss => ss.ParentRequestId != null && requestIds.Contains((long)ss.ParentRequestId) && years.Contains(ss.YearOfRest.Year) && ss.StatusId == (long)StatusEnum.CertificateIssued).ToList();


							var resultRequests = initialRequests.Join(descendantRequests, r => r.Id, dr => dr.ParentRequestId, (r, dr) => new Request { Id = dr.Id, Tour = dr.Tour, TourId = r.TourId, TypeOfRest = dr.TypeOfRest, RequestOnMoney = r.RequestOnMoney, TypeOfRestId = r.TypeOfRestId, YearOfRest = r.YearOfRest }).ToList();
							resultRequests.AddRange(requests.Where(req => !resultRequests.Any(rr => rr.Id == req.Id) && req.StatusId == (long?)StatusEnum.CertificateIssued));


							if (_FIRST_LINE)
							{
								PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);


								PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Информация об услугах, оказанных ребёнку/детям в течение последних 3-х лет", _HEADER_FONT));

								_FIRST_LINE = false;
							}

							PdfAddParagraph(document, 0, 10, Element.ALIGN_LEFT, 0,
								new Chunk($"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}, {child.Snils}", _MAIN_TXT_FONT));


							GetPdfTable(resultRequests, document, years);
						}
					}

					document.Close();

				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в предоставлении услуг отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse108013(Request request, Account account)

		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					var simple = false;
					if (request.StatusId == (long)StatusEnum.CertificateIssued)
					{
						PdfAddParagraph(document,
							"Уведомление о неучастии заявителя в выборе конкретной организации отдыха и оздоровления",
							0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						simple = true;
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
							new Chunk("Уведомление об отказе в предоставлении услуг отдыха и оздоровления", _TITLE_FONT),
							Chunk.NEWLINE,
							new Chunk("в связи с неучастием заявителя во втором этапе заявочной кампании", _TITLE_FONT));
					}

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));
					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
								new Chunk(
									simple ? "Данные ребенка/детей: " : "Данные ребенка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ",
									_HEADER_FONT),
								new Chunk(
									$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
									_MAIN_TXT_ITALIC_FONT));
							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
								new Chunk("Льготная категория: ", _HEADER_FONT),
								new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
						}
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребенка (детей) /лица из числа детей-сирот и детей, оставшихся без попечения родителей: ",
								_HEADER_FONT),
							new Chunk(
								$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk(
								"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет",
								_MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk(simple ? "Услуга оказана. Решение отрицательное." : "отказ в предоставлении услуг отдыха и оздоровления.",
							_MAIN_TXT_ITALIC_FONT));
					if (simple)
					{
						//PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						//    new Chunk("Основание: ", _HEADER_FONT),
						//        new Chunk(("пункт 8.8. Порядка организации отдыха и оздоровления детей, находящихся"), _MAIN_TXT_FONT), Chunk.NEWLINE,
						//            new Chunk(("в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы"), _MAIN_TXT_FONT), Chunk.NEWLINE,
						//                new Chunk(("от 22 февраля 2017 г. №56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\"") +
						//                           " (далее – Порядок).", _MAIN_TXT_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
							new Chunk("Основание: ", _HEADER_FONT),
								new Chunk("пункт 8.8. Порядка организации отдыха и оздоровления детей, находящихся " +
											"в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы " +
											"от 22 февраля 2017 г. №\u00A056-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\"" +
											 " (далее – Порядок).", _MAIN_TXT_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED_ALL, 0, new Chunk("Основание отказа: ", _HEADER_FONT),
							new Chunk(
								("пункты 8.9. и 9.1.7 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы"
								), _MAIN_TXT_FONT)
						);

						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
							new Chunk(
								("от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\""
								) + " (далее – Порядок).", _MAIN_TXT_FONT)
							);
					}
					if (simple)
					{
						//PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("          ", _HEADER_FONT),
						//    new Chunk(("В случае неучастия заявителя, подавшего заявление о предоставлении услуг отдыха"
						//        ), _MAIN_TXT_FONT),
						//    Chunk.NEWLINE,
						//    new Chunk(("и оздоровления в отношении организации индивидуального выездного отдыха либо совместного выездного отдыха ребенка, отнесенного к одной из категорий детей, находящихся в трудной жизненной ситуации и указанных в пунктах 3.1.3 и 3.1.4 Порядка, заявителя, подавшего заявление"
						//        ), _MAIN_TXT_FONT),
						//    Chunk.NEWLINE,
						//    new Chunk(("о предоставлении услуг отдыха и оздоровления в отношении организации индивидуального выездного отдыха ребенка, отнесенного к одной из категорий детей, находящихся в трудной жизненной ситуации"
						//        ), _MAIN_TXT_FONT),
						//    Chunk.NEWLINE,
						//    new Chunk(("и указанных в пунктах 3.1.5-3.1.13 Порядка, во втором этапе заявочной кампании, услуга отдыха и оздоровления считается предоставленной."
						//        ), _MAIN_TXT_FONT));
						PdfAddParagraph(document, 1f, 1f, Element.ALIGN_JUSTIFIED, 0, Chunk.TABBING,
							new Chunk($"В случае неучастия заявителя, подавшего заявление о предоставлении услуг отдыха " +
										"и оздоровления в отношении организации индивидуального выездного отдыха либо совместного выездного отдыха ребенка, отнесенного к одной из категорий детей, находящихся в трудной жизненной ситуации и указанных в пунктах 3.1.3 и 3.1.4 Порядка, заявителя, подавшего заявление " +
										"о предоставлении услуг отдыха и оздоровления в отношении организации индивидуального выездного отдыха ребенка, отнесенного к одной из категорий детей, находящихся в трудной жизненной ситуации " +
										"и указанных в пунктах 3.1.5-3.1.13 Порядка, во втором этапе заявочной кампании, услуга отдыха и оздоровления считается предоставленной.", _MAIN_TXT_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED_ALL, 25,
							new Chunk(("В случае неучастия заявителя, подавшего заявление о предоставлении услуг отдыха"
								), _MAIN_TXT_FONT));

						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
							new Chunk(("и оздоровления в отношении организации индивидуального выездного отдыха либо совместного выездного отдыха ребенка, находящегося в трудной жизненной ситуации и указанного в пункте 3.1.2 Порядка, лица из числа детей-сирот и детей, оставшихся без попечения родителей, во втором этапе заявочной кампании, услуга отдыха и оздоровления не считается оказанной."
								), _MAIN_TXT_FONT));
					}

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление в связи с неучастием" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuseCompensation(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var isCompensationYouth = request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest;

					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					//PdfAddParagraph(document, "", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
					PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
						new Chunk("Уведомление об отказе в выплате компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления", _TITLE_FONT));


					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата обращения: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));


					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
					   new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
					   new Chunk("отказ в предоставлении выплаты компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления.", _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
					   new Chunk("Основание: ", _HEADER_FONT),
					   new Chunk("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее – Порядок).", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
					  new Chunk("Причина отказа в выплате: ", _HEADER_FONT),
					  new Chunk(request.DeclineReason?.Name, _MAIN_TXT_FONT));


					//SignWorkerBlock(document, account);
					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuseCompensationYouthRest(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var isCompensationYouth = request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest;

					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					//PdfAddParagraph(document, "", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
					PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
						new Chunk("Уведомление об отказе в выплате компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления", _TITLE_FONT));


					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата обращения: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));


					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
					   new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
					   new Chunk("отказ в предоставлении выплаты компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления.", _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
					   new Chunk("Основание: ", _HEADER_FONT),
					   new Chunk("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее – Порядок).", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
					  new Chunk("Причина отказа в выплате: ", _HEADER_FONT),
					  new Chunk(request.DeclineReason?.Name, _MAIN_TXT_FONT));


					//SignWorkerBlock(document, account);
					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse1080(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, "Уведомление об отказе в предоставлении услуг отдыха и оздоровления", 0,
						1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));

					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребенка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ",
								_HEADER_FONT),
							new Chunk("\n", _MAIN_TXT_ITALIC_FONT),
							new Chunk(
								$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ",
								_HEADER_FONT),
							new Chunk("\n", _MAIN_TXT_ITALIC_FONT),
							new Chunk(
								$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk(
								"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет",
								_MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("отказ в предоставлении услуг отдыха и оздоровления.", _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Основание отказа: ", _HEADER_FONT),
						new Chunk(_FEDERAL_LAW, _MAIN_TXT_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Причина отказа: ", _HEADER_FONT),
						new Chunk(request.DeclineReason?.Name, _MAIN_TXT_FONT));

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в предоставлении услуг отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		//NotificationRefuse108013
		protected override IDocument NotificationRefuseContent(Request request, Account account)
		{
			return NotificationRefuse108013(request, account);
		}
		#endregion

		#region NotificationAboutDecision

		protected override IDocument NotificationAboutCompensationIssued(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var isCompensationYouth = request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest;

					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					if (isCompensationYouth)
					{
						//PdfAddParagraph(document, "", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
						PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
							new Chunk("Уведомление о предоставлении выплаты компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления", _TITLE_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
							new Chunk("Уведомление о предоставлении выплаты компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления", _TITLE_FONT));
					}

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));

					if (request.Child != null)
					{
						foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
								new Chunk(
									$"{(isCompensationYouth ? "Данные лица из числа детей-сирот и детей, оставшихся без попечения родителей" : "Данные о ребёнке")}: ",
									_HEADER_FONT),
								new Chunk(
									$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
									_MAIN_TXT_ITALIC_FONT));
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
								new Chunk($"{(isCompensationYouth ? "Льготная категория" : "Льготная категория ребёнка")}: ", _HEADER_FONT),
								new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
						}
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
					   new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
					   new Chunk("Услуга оказана.", _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
					   new Chunk("Основание: ", _HEADER_FONT),
					   new Chunk("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\".", _MAIN_TXT_FONT));


					//SignWorkerBlock(document, account);
					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о предоставлении выплаты компенсации за самостоятельно приобретенную путевку" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationAboutCertificate(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, "Уведомление о предоставлении сертификата на отдых и оздоровление", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name ?? "сертификат на отдых и оздоровление", _MAIN_TXT_ITALIC_FONT));

					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
								new Chunk(
									"Данные лиц, указанных в сертификате: ",
									_HEADER_FONT),
								new Chunk(
									$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
									_MAIN_TXT_ITALIC_FONT));
							PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
								new Chunk("Льготная категория ребенка: ", _HEADER_FONT),
								new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
						}
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn18 ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ",
								_HEADER_FONT),
							new Chunk(
								$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk(
								"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет",
								_MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("Услуга оказана.", _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Номер сертификата: ", _HEADER_FONT),
						new Chunk(request.CertificateNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));


					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Основание: ", _HEADER_FONT),
						new Chunk($"{_FEDERAL_SHORT2021_LAW}.", _MAIN_TXT_ITALIC_FONT));

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о предоставлении сертификата на отдых и оздоровление" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationAboutTour(Request request, Account account, bool forPrint = false)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document,
						"Уведомление о предоставлении бесплатной путевки для отдыха",
						0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document,
						"и оздоровления",
						0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Вид отдыха: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));

					if (request.Child != null && request.Child.Count > 0)
					{
						if (request.Applicant.IsAccomp || request.Attendant.Any())
						{
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
										new Chunk(
											"Данные лиц, указанных в путевке: ",
											_HEADER_FONT));
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
										new Chunk(
											"Сопровождающие: ",
											_HEADER_FONT));
							if (request.Applicant.IsAccomp)
							{
								PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
										new Chunk(
											$"{request.Applicant.LastName} {request.Applicant.FirstName} {request.Applicant.MiddleName}, {request.Applicant.DateOfBirth.FormatExGR(string.Empty)}",
											_MAIN_TXT_ITALIC_FONT));
							}
							if (request.Attendant.Any())
							{
								foreach (var attendant in request.Attendant.Where(c => !c.IsDeleted))
								{
									PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
											new Chunk(
												$"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}, {attendant.DateOfBirth.FormatExGR(string.Empty)}",
												_MAIN_TXT_ITALIC_FONT));
								}
							}
							PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
									new Chunk(
										"Дети: ",
										_HEADER_FONT));
						}

						{
							foreach (var child in request.Child.Where(c => !c.IsDeleted))
							{
								if (request.Child.Count == 1 && request.TypeOfRest.ParentId != 2 && request.TypeOfRest.ParentId != 9)
								{
									PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
										new Chunk(
											"Данные лиц, указанных в путевке: ",
											_HEADER_FONT),
										new Chunk(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
											_MAIN_TXT_ITALIC_FONT));
									PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
										new Chunk("Льготная категория: ", _HEADER_FONT),
										new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
								}
								else
								{
									PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
										//new Chunk(
										//    "Данные лиц, указанных в путевке: ",
										//    _HEADER_FONT),
										new Chunk(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}",
											_MAIN_TXT_ITALIC_FONT));
									PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
										new Chunk("Льготная категория: ", _HEADER_FONT),
										new Chunk($"{child.BenefitType?.Name}", _MAIN_TXT_ITALIC_FONT));
								}
							}
						}
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk(
								"Данные лиц, указанных в путевке: ",
								_HEADER_FONT),
							new Chunk(
								$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
								_MAIN_TXT_ITALIC_FONT));
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
							new Chunk("Льготная категория: ", _HEADER_FONT),
							new Chunk(
								"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет",
								_MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Результат рассмотрения заявления: ", _HEADER_FONT),
						new Chunk("Услуга оказана.", _MAIN_TXT_ITALIC_FONT));
					/*PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
                        new Chunk("Дата и время выбора варианта организации отдыха и оздоровления: ", _HEADER_FONT),
                        new Chunk($"{request.DateChangeStatus:dd.MM.yyyy HH:mm}", _MAIN_TXT_ITALIC_FONT));*/
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер путевки: ", _HEADER_FONT),
						new Chunk(request.CertificateNumber, _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Организация отдыха и оздоровления: ", _HEADER_FONT),
						new Chunk(request.Tour?.Hotels?.Name, _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Время отдыха: ", _HEADER_FONT),
						new Chunk(
							$"{request.TimeOfRest?.Name} ({request.Tour?.DateIncome.FormatEx()} - {request.Tour?.DateOutcome.FormatEx()})",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0, new Chunk("Основание: ", _HEADER_FONT),
						new Chunk($"{_FEDERAL_SHORT2021_LAW}.", _MAIN_TXT_ITALIC_FONT));

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о предоставлении бесплатной путевки для отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}
		#endregion

		public override IDocument NotificationOrgChoose(Request request, Account account)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);
					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, "Уведомление о необходимости выбора организации отдыха и оздоровления", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0, new Chunk($"Уважаемый(ая) {applicant.LastName} {applicant.FirstName} {applicant.MiddleName},", _MAIN_TXT_FONT));
					PdfAddParagraph(document, $"Ваше заявление от {request.DateRequest.FormatEx("dd.MM.yyyy")} г. № {request.RequestNumber} о предоставлении услуг отдыха и оздоровления (далее – заявление) рассмотрено.", 0, 1, Element.ALIGN_JUSTIFIED, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("В период второго этапа заявочной кампании (с 7 по 21 февраля 2023 г.) Вам необходимо дополнить Ваше заявление сведениями о конкретной организации отдыха и оздоровления. Выбор конкретной организации отдыха и оздоровления осуществляется из числа предлагаемых ГАУК\u00A0\"МОСГОРТУР\" в соответствии с указанными Вами на первом этапе заявочной кампании сведениями о приоритетном времени, направлении отдыха и оздоровления и количестве детей.", _MAIN_TXT_FONT)); ;

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("В случае подачи заявления через подсистему \"личный кабинет\" Портала Мэра и Правительства Москвы mos.ru (далее – подсистема \"личный кабинет\" Портала), дополнение формы заявления сведениями о конкретной организации отдыха и оздоровления осуществляется в подсистеме \"личный кабинет\" Портала.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("В случае подачи заявления при личном обращении в офисе ГАУК \"МОСГОРТУР\", дополнение формы заявления сведениями о конкретной организации отдыха и оздоровления возможно только при личном обращении заявителя в офис ГАУК \"МОСГОРТУР\" по адресу: г.\u00A0Москва, Малый Харитоньевский переулок д. 6, стр. 3.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("Прием в офисе ГАУК \"МОСГОРТУР\" осуществляется исключительно по предварительной записи.", _HEADER_FONT),
						new Chunk(" Запись производится через Портал Мэра и Правительства Москвы mos.ru  (далее – Портал) или при личном визите заявителя в офис ГАУК \"МОСГОРТУР\". Запись производится на свободную дату и время.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("В случае если на втором этапе заявочной кампании Вас не устроит ни один из вариантов организаций отдыха и оздоровления, предлагаемых ГАУК \"МОСГОРТУР\", на втором этапе заявочной кампании в период с 7 по 21 февраля 2023 г. Вам необходимо отказаться от всех предложенных организаций отдыха и оздоровления.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("При подаче заявления через подсистему \"личный кабинет\" Портала для отказа от всех предложенных организаций отдыха и оздоровления необходимо нажать соответствующую \"галочку\" в интерактивном поле Портала: \"Я отказываюсь от предложенных вариантов организаций отдыха и оздоровления\".", _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL, new Chunk("При подаче заявления в офисе ГАУК \"МОСГОРТУР\" для отказа от всех предложенных организаций отдыха и оздоровления необходимо лично обратиться в офис ГАУК \"МОСГОРТУР\".", _MAIN_TXT_FONT));

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о необходимости выбора организации отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		///     Уведомление о замене сопровождающего
		/// </summary>
		protected override IDocument InnerNotificationAttendantChange(Request request, Account account, long oldAttendantId, long attendantId, string changeDate = null)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };
			Applicant oldAttendant = null;
			Applicant attendant = request.Attendant.FirstOrDefault(att => att.Id == attendantId);
			if (request.Applicant.Id == oldAttendantId)
				oldAttendant = applicant;
			else oldAttendant = request.Attendant.FirstOrDefault(att => att.Id == oldAttendantId);
			var isCert = request.TypeOfRestId == (long)TypeOfRestEnum.Money || request.TypeOfRest.ParentId == (long)TypeOfRestEnum.Money;
			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);


					//PdfAddParagraph(document, "Уведомление о регистрации заявления о предоставлении услуг отдыха и оздоровления", 0, 1, Element.ALIGN_CENTER, 0, _TITLE_FONT);
					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
							new Chunk("Уведомление о замене сопровождающего лица в предоставленном сертификате на отдых и оздоровление", _TITLE_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
						   new Chunk("Уведомление о замене сопровождающего лица в предоставленной бесплатной путевке для отдыха и оздоровления", _TITLE_FONT));
					}

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);


					//PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата и время регистрации заявления: ", _HEADER_FONT),
					//    new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					//PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Номер заявления: ", _HEADER_FONT),
					//    new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные заявителя: ", _HEADER_FONT),
						new Chunk(
							$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));
					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk(applicant.Phone + ", " + applicant.Email, _MAIN_TXT_ITALIC_FONT));
					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Номер сертификата: ", _HEADER_FONT),
						new Chunk(request.CertificateNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
					   new Chunk("Номер путевки: ", _HEADER_FONT),
					   new Chunk(request.CertificateNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));
					}

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
					new Chunk("Номер заявления: ", _HEADER_FONT),
					new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					if (!request.Tour.IsNullOrEmpty())
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
					   new Chunk("Организация отдыха и оздоровления: ", _HEADER_FONT),
					   new Chunk(request.Tour.Hotels.NameOrganization ?? request.Tour.Hotels.Name, _MAIN_TXT_ITALIC_FONT));

					if (!request.Tour.IsNullOrEmpty())
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Время отдыха: ", _HEADER_FONT),
						new Chunk(request.Tour.DateIncome.FormatEx("dd.MM.yyyy") + " - " + request.Tour.DateOutcome.FormatEx("dd.MM.yyyy"), _MAIN_TXT_ITALIC_FONT));


					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные сопровождающего лица, указанного при подаче заявления о предоставлении услуг отдыха и оздоровления: ", _HEADER_FONT),
						new Chunk(
							$"{oldAttendant.LastName} {oldAttendant.FirstName} {oldAttendant.MiddleName}, {oldAttendant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Данные предлагаемого к замене сопровождающего лица, указанного в заявлении о замене сопровождающего лица: ", _HEADER_FONT),
						new Chunk(
							$"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}, {attendant.DateOfBirth.FormatExGR(string.Empty)}",
							_MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Основание приема заявления: ", _HEADER_FONT),
					   new Chunk("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\".", _MAIN_TXT_FONT));

					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения: ", _HEADER_FONT),
						new Chunk("Ваше заявление о замене сопровождающего лица в предоставленном сертификате на отдых и оздоровление рассмотрено. Сопровождающее лицо изменено. Сертификат на отдых и оздоровление с измененными сведениями о сопровождающем лице прилагается.", _MAIN_TXT_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Результат рассмотрения: ", _HEADER_FONT),
						new Chunk("Ваше заявление о замене сопровождающего лица в предоставленной путевке для отдыха и оздоровления рассмотрено положительно. Путевка для отдыха и оздоровления с измененными сведениями о сопровождающем лице прилагается.", _MAIN_TXT_FONT));

					}
					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата написания заявления о замене сопровождающего лица в предоставленном сертификате на отдых и оздоровление: ", _HEADER_FONT),
						new Chunk(oldAttendant.AttendantChangeRequestDate.HasValue ? oldAttendant.AttendantChangeRequestDate.FormatEx("dd.MM.yyyy") : DateTime.Now.FormatEx("dd.MM.yyyy"), _MAIN_TXT_ITALIC_FONT));
						//PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата замены сопровождающего лица в предоставленном сертификате на отдых и оздоровление: ", _HEADER_FONT),
						//new Chunk(oldAttendant.AttendantChangeRequestDate.HasValue ? oldAttendant.AttendantChangeDate.FormatEx("dd.MM.yyyy") : DateTime.Now.FormatEx("dd.MM.yyyy"), _MAIN_TXT_ITALIC_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата написания заявления о замене сопровождающего лица в предоставленной бесплатной путевке для отдыха и оздоровления: ", _HEADER_FONT),
						new Chunk(oldAttendant.AttendantChangeRequestDate.HasValue ? oldAttendant.AttendantChangeRequestDate.FormatEx("dd.MM.yyyy") : DateTime.Now.FormatEx("dd.MM.yyyy"), _MAIN_TXT_ITALIC_FONT));
						//PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0, new Chunk("Дата замены сопровождающего лица в предоставленной бесплатной путевке для отдыха и оздоровления: ", _HEADER_FONT),
						//new Chunk(oldAttendant.AttendantChangeRequestDate.HasValue ? oldAttendant.AttendantChangeDate.FormatEx("dd.MM.yyyy") : DateTime.Now.FormatEx("dd.MM.yyyy"), _MAIN_TXT_ITALIC_FONT));
					}

					//PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
					//   new Chunk("Путевка для отдыха и оздоровления с измененными сведениями о сопровождающем лице прилагается.", _MAIN_TXT_FONT));

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					if (isCert)
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						   new Chunk("Приложение: сертификат на отдых и оздоровление.", _MAIN_TXT_FONT));
					}
					else
					{
						PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, _FIRST_LINE_SMALL,
						   new Chunk("Приложение: путевка для отдыха и оздоровления.", _MAIN_TXT_FONT));
					}

					//SignWorkerBlock(document, account);
					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о замене сопровождающего лица" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		public override IDocument NotificationDeclineRegistryReg(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var document = new Document(PageSize.A4, 60, 40, 40, 40))
				{
					var writer = PdfWriter.GetInstance(document, ms);

					document.Open();
					writer.SpaceCharRatio = PdfWriter.NO_SPACE_CHAR_RATIO;

					PdfAddHeader(document);

					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);
					PdfAddParagraph(document, _SPACE, 0, 1, Element.ALIGN_LEFT, 0, _MAIN_TXT_FONT);

					PdfAddParagraph(document, 0, 1, Element.ALIGN_CENTER, 0,
						new Chunk("Уведомление об отказе в регистрации заявления о предоставлении услуг отдыха и оздоровления", _TITLE_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Номер отказа: ", _HEADER_FONT),
						new Chunk(request.RequestNumber.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Дата обращения: ", _HEADER_FONT),
						new Chunk(request.DateRequest.FormatEx(), _MAIN_TXT_ITALIC_FONT));

					string applicantBirthDate = request.Applicant?.DateOfBirth?.ToString("dd.MM.yyyy");
					string applicantData = $"{request.Applicant?.LastName ?? string.Empty} {request.Applicant?.FirstName ?? string.Empty} {request.Applicant?.MiddleName ?? string.Empty}";
					applicantData += !string.IsNullOrWhiteSpace(applicantBirthDate) ? $", {applicantBirthDate} г.р." : string.Empty;

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
					   new Chunk("Данные обратившегося лица: ", _HEADER_FONT),
					   new Chunk(applicantData, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Контактная информация: ", _HEADER_FONT),
						new Chunk((request.Applicant?.Phone ?? "-")
									+ ", " + (request.Applicant?.Email ?? "-")
									, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Цель обращения: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Parent?.Name, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_LEFT, 0,
						new Chunk("Вид услуги: ", _HEADER_FONT),
						new Chunk(request.TypeOfRest?.Name, _MAIN_TXT_ITALIC_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Основание отказа в регистрации заявления: ", _HEADER_FONT),
						new Chunk(_FEDERAL_LAW, _MAIN_TXT_FONT));

					PdfAddParagraph(document, 0, 1, Element.ALIGN_JUSTIFIED, 0,
						new Chunk("Причина отказа в регистрации заявления: ", _HEADER_FONT),
						new Chunk(request.DeclineReason.Name, _MAIN_TXT_FONT));

					document.Close();
				}

				return new CshedDocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в регистрации заявления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		#endregion

		#region PublicMethods

		public ICshedDocument NotificationEcpSigned(IDocument document, CertificateToApply certificate, string cshedGuid, string cshedSignGuid, string documentlink, string signlink, string[] certificateDates)
		{
			using (var templateStream = new MemoryStream(document.FileBody))
			{
				using (var streamTemplate = new MemoryStream())
				{

					templateStream.Seek(0, SeekOrigin.Begin);
					using (var readerTemplate = new PdfReader(templateStream))
					{
						using (var pdfStamper = new PdfStamper(readerTemplate, streamTemplate, '1'))
						{
							var page = pdfStamper.Reader.NumberOfPages + 1;
							var pageSize = pdfStamper.Reader.GetPageSizeWithRotation(1);
							pdfStamper.InsertPage(page, pageSize);

							var over = pdfStamper.GetOverContent(page);

							float currentPosY = pageSize.Height - 60;
							float currentPosX = 70;
							float lineOffset = 14;
							float fontSize = 11;
							float titleFontSize = 14;

							Rectangle linkDocumentRect; // Область для линка "Ссылка для скачивания оригинала документа"
							Rectangle linkSignRect; // Область для линка "Ссылка для скачивания электронной подписи оригинала документа"

							var font = _ECP_BASE_FONT;
							var boldFont = _ECP_BASE_BOLD_FONT;

							var linkDocument = documentlink.Replace("{0}", cshedGuid); // Ссылка для скачивания оригинала документа
							var linkSign = signlink.Replace("{0}", cshedGuid).Replace("{1}", cshedSignGuid); // Ссылка для скачивания электронной подписи оригинала документа

							over.BeginText();

							//Сведения о документе
							over.SetFontAndSize(boldFont, titleFontSize);
							over.ShowTextAligned(Element.ALIGN_LEFT, "Сведения о документе", currentPosX, currentPosY -= lineOffset, 0);

							over.SetFontAndSize(font, fontSize);
							over.ShowTextAligned(Element.ALIGN_LEFT, "Место хранения оригинала документа, подписанного электронной подписью:", currentPosX, currentPosY -= lineOffset, 0);
							over.ShowTextAligned(Element.ALIGN_LEFT, "Центральное хранилище электронных документов Правительства Москвы", currentPosX + 30, currentPosY -= lineOffset, 0);

							currentPosY -= lineOffset;
							linkDocumentRect = new Rectangle(currentPosX, currentPosY, currentPosX + 233, currentPosY + 7);

							over.SetColorFill(BaseColor.BLUE);
							over.ShowTextAligned(Element.ALIGN_LEFT, "Ссылка для скачивания оригинала документа", currentPosX, currentPosY, 0);
							over.SetColorFill(BaseColor.BLACK);

							//Сведения об ЭП № 1
							over.SetFontAndSize(boldFont, titleFontSize);
							over.ShowTextAligned(Element.ALIGN_LEFT, "Сведения об ЭП № 1", currentPosX, currentPosY -= lineOffset + 10, 0);

							over.SetFontAndSize(font, fontSize);
							over.ShowTextAligned(Element.ALIGN_LEFT, "Организация:", currentPosX, currentPosY -= lineOffset, 0);
							over.ShowTextAligned(Element.ALIGN_LEFT, "ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ УЧРЕЖДЕНИЕ КУЛЬТУРЫ ГОРОДА МОСКВЫ", currentPosX + 30, currentPosY -= lineOffset, 0);
							over.ShowTextAligned(Element.ALIGN_LEFT, "\"МОСКОВСКОЕ АГЕНТСТВО ОРГАНИЗАЦИИ ОТДЫХА И ТУРИЗМА\"", currentPosX + 30, currentPosY -= lineOffset, 0);

							over.ShowTextAligned(Element.ALIGN_LEFT, "Должность подписанта:", currentPosX, currentPosY -= lineOffset, 0);
							var accountPosition = string.IsNullOrEmpty(certificate.Account.Position) ? "Сведения в сертификате электронной подписи отсутствуют" : certificate.Account.Position;
							over.ShowTextAligned(Element.ALIGN_LEFT, accountPosition, currentPosX + 30, currentPosY -= lineOffset, 0);

							over.ShowTextAligned(Element.ALIGN_LEFT, "Подписант:", currentPosX, currentPosY -= lineOffset, 0);
							over.ShowTextAligned(Element.ALIGN_LEFT, certificate.Account.Name, currentPosX + 30, currentPosY -= lineOffset, 0);

							over.ShowTextAligned(Element.ALIGN_LEFT, "Дата и время подписания документа:", currentPosX, currentPosY -= lineOffset, 0);
							over.ShowTextAligned(Element.ALIGN_LEFT, DateTime.Now.FormatEx("dd.MM.yyyy HH:mm:ss"), currentPosX + 30, currentPosY -= lineOffset, 0);

							currentPosY -= lineOffset;
							linkSignRect = new Rectangle(currentPosX, currentPosY, currentPosX + 345, currentPosY + 7);

							over.SetColorFill(BaseColor.BLUE);
							over.ShowTextAligned(Element.ALIGN_LEFT, "Ссылка для скачивания электронной подписи оригинала документа", currentPosX, currentPosY, 0);
							over.SetColorFill(BaseColor.BLACK);

							// Рамка с данными ЭЦП
							currentPosY -= 30;
							var framePosY = currentPosY + 10;
							var firstColumnOffsetPosX = 3; // Смещение первого столбца в рамке
							var secondColumnOffsetPosX = 70; // Смещение второго столбца в рамке
							var frameTitleOffsetPosX = 140;
							var frameWidth = 300;
							var frameHeight = 83;
							var imageHeight = 25; // Высота иконки, ширина рассчитывается автоматически
							fontSize = 9;
							lineOffset = 12;

							over.SetFontAndSize(boldFont, fontSize);
							over.ShowTextAligned(Element.ALIGN_CENTER, "ДОКУМЕНТ ПОДПИСАН", currentPosX + frameTitleOffsetPosX, currentPosY, 0);
							over.ShowTextAligned(Element.ALIGN_CENTER, "ЭЛЕКТРОННОЙ ПОДПИСЬЮ", currentPosX + frameTitleOffsetPosX, currentPosY -= lineOffset, 0);

							over.SetFontAndSize(font, fontSize);
							over.ShowTextAligned(Element.ALIGN_LEFT, "Сертификат:", currentPosX + firstColumnOffsetPosX, currentPosY -= 20, 0);
							over.ShowTextAligned(Element.ALIGN_LEFT, certificate.CertificateKey.Replace(" ", "").Trim(), currentPosX + secondColumnOffsetPosX, currentPosY, 0);

							over.ShowTextAligned(Element.ALIGN_LEFT, "Владелец:", currentPosX + firstColumnOffsetPosX, currentPosY -= lineOffset, 0);
							over.SetFontAndSize(boldFont, fontSize);
							over.ShowTextAligned(Element.ALIGN_LEFT, certificate.Account.Name, currentPosX + secondColumnOffsetPosX, currentPosY, 0);

							over.SetFontAndSize(font, fontSize);
							over.ShowTextAligned(Element.ALIGN_LEFT, "Действителен:", currentPosX + firstColumnOffsetPosX, currentPosY -= lineOffset, 0);
							if (certificateDates != null && certificateDates.Length >= 2)
							{
								over.ShowTextAligned(Element.ALIGN_LEFT, $"с {certificateDates[0].Substring(0, 10)} по {certificateDates[1].Substring(0, 10)}", currentPosX + secondColumnOffsetPosX, currentPosY, 0);
							}

							over.SetFontAndSize(boldFont, fontSize);
							over.ShowTextAligned(Element.ALIGN_LEFT, "ПРОВЕРКА:", currentPosX + firstColumnOffsetPosX, currentPosY -= lineOffset, 0);
							over.SetFontAndSize(_BASE_FONT, fontSize);
							over.ShowTextAligned(Element.ALIGN_LEFT, "ЭП подтверждена", currentPosX + secondColumnOffsetPosX, currentPosY, 0);

							over.EndText();

							// Иконка
							if (File.Exists(_ECP_ICON_PATH))
							{
								var ecpIcon = Image.GetInstance(_ECP_ICON_PATH);
								var imageAspectRatio = ecpIcon.Height / ecpIcon.Width;
								var imageWidth = imageHeight / imageAspectRatio;

								ecpIcon.ScaleAbsoluteWidth(imageWidth);
								ecpIcon.ScaleAbsoluteHeight(imageHeight);
								ecpIcon.SetAbsolutePosition(currentPosX + 6, framePosY - imageHeight - 3);

								over.AddImage(ecpIcon);
							}

							// Рамка
							Rectangle rect = new Rectangle(currentPosX, framePosY - frameHeight, currentPosX + frameWidth, framePosY);
							rect.Border = Rectangle.BOX;
							rect.BorderColor = BaseColor.BLACK;
							rect.BorderWidth = 1;
							over.Rectangle(rect);

							// Добавляем ссылки
							over.Rectangle(linkDocumentRect);
							PdfAnnotation linkDocumentAnnotation = PdfAnnotation.CreateLink(
								pdfStamper.Writer, linkDocumentRect, PdfAnnotation.HIGHLIGHT_INVERT,
								new PdfAction(linkDocument)
							);
							pdfStamper.AddAnnotation(linkDocumentAnnotation, page);


							over.Rectangle(linkSignRect);
							PdfAnnotation linkSignAnnotation = PdfAnnotation.CreateLink(
								pdfStamper.Writer, linkSignRect, PdfAnnotation.HIGHLIGHT_INVERT,
								new PdfAction(linkSign)
							);
							pdfStamper.AddAnnotation(linkSignAnnotation, page);
						}
					}
					return new CshedDocumentResult
					{
						FileBody = streamTemplate.ToArray(),
						FileName = document.FileName,
						MimeType = MimeType,
						MimeTypeShort = Extension
					};
				}
			}
		}

		#endregion

		#region PrivateMethods

		/// <summary>
		/// сертификат на отдых 1 ребенок
		/// </summary>
		private string CertificateOnMoney(Request request, Stream newStream)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			var assembly = Assembly.Load("RestChild.Templates");

			var requestChildren = request.Child.Where(c => !c.IsDeleted).ToList();

			using (var templateStream = new MemoryStream())
			{
				using (var streamTemplate = assembly.GetManifestResourceStream("RestChild.Templates.singlePayment2021.pdf"))
				{
					using (var readerTemplate = new PdfReader(streamTemplate))
					{
						using (var doc = new Document())
						{
							using (var copy = new PdfCopy(doc, templateStream))
							{
								doc.Open();
								var childrenCount = requestChildren.Count;
								while (childrenCount > 0)
								{
									copy.AddPage(copy.GetImportedPage(readerTemplate, 1));
									childrenCount--;
								}

								if (request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn18)
								{
									copy.AddPage(copy.GetImportedPage(readerTemplate, 1));
								}

								copy.AddPage(copy.GetImportedPage(readerTemplate, 2));
								copy.CloseStream = false;
								doc.Close();
							}
						}
					}

					templateStream.Seek(0, SeekOrigin.Begin);
					using (var readerTemplate = new PdfReader(templateStream))
					{
						using (var pdfStamper = new PdfStamper(readerTemplate, newStream, '1'))
						{
							var page = 1;
							if (request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn18)
							{
								var over = pdfStamper.GetOverContent(page);
								over.BeginText();
								over.SetFontAndSize(_customFont, 14);
								over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateDate?.ToLongDateString() ?? request.DateChangeStatus?.ToLongDateString() ?? DateTime.Now.ToLongDateString(), 280, 458, 0);
								over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateNumber.FormatEx(), 475, 458, 0);
								over.SetFontAndSize(_customFont, 12);

								WriteByTable(over, _FONT_12, 270, 400, 500, $"{applicant.LastName}  {applicant.FirstName}  {applicant.MiddleName}".Trim());

								WriteByTable(over, _FONT_12, 40, 366, 700, $"{applicant.DateOfBirth.FormatEx()},  {applicant.DocumentType.Name},  {applicant.DocumentSeria},  {applicant.DocumentNumber}".Trim(), 1);

								WriteByTable(over, _FONT_12, 270, 243, 500, $"{request.RequestNumber}, {applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DocumentType.Name}, {applicant.DocumentSeria} {applicant.DocumentNumber}".Trim());
								over.EndText();
							}
							else
							{
								foreach (var child in requestChildren)
								{
									var over = pdfStamper.GetOverContent(page);
									if (page == 1) page++;
									page++;
									over.BeginText();
									over.SetFontAndSize(_customFont, 14);
									over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateDate?.ToLongDateString() ?? request.DateChangeStatus?.ToLongDateString() ?? DateTime.Now.ToLongDateString(), 280, 458, 0);
									over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateNumber.FormatEx(), 475, 458, 0);

									WriteByTable(over, _FONT_12, 270, 400, 500, $"{child.LastName}  {child.FirstName}  {child.MiddleName}".Trim());

									WriteByTable(over, _FONT_12, 40, 366, 700, $"{child.DateOfBirth.FormatEx()},  {child.DocumentType.Name},  {child.DocumentSeria},  {child.DocumentNumber}".Trim(), 1);

									if (applicant.IsAccomp || request.Attendant.Any(a => a.IsAccomp && !a.IsDeleted))
									{
										var attendant =
											(applicant.IsAccomp && !applicant.IsDeleted
												? applicant
												: request.Attendant.FirstOrDefault(a => a.IsAccomp && !a.IsDeleted)) ??
											new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

										WriteByTable(over, _FONT_12, 360, 333, 400, $"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}".Trim());

										WriteByTable(over, _FONT_12, 40, 293, 700, $"{attendant.DateOfBirth.FormatEx()}, {attendant.DocumentType.Name}, {attendant.DocumentSeria}, {attendant.DocumentNumber}".Trim(), 1);

									}

									WriteByTable(over, _FONT_12, 270, 243, 500, $"{request.RequestNumber}, {applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DocumentType.Name}, {applicant.DocumentSeria} {applicant.DocumentNumber}".Trim());

									over.EndText();
								}
							}
						}
					}
				}
			}
			return "Сертификат.pdf";
		}

		/// <summary>
		/// сертификат на отдых много детей
		/// </summary>
		private string CertificateOnMoneyMulti(Request request, Stream newStream)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			var assembly = Assembly.Load("RestChild.Templates");

			using (var templateStream = new MemoryStream())
			{
				using (var streamTemplate = assembly.GetManifestResourceStream("RestChild.Templates.multiPayment.pdf"))
				{
					streamTemplate?.CopyTo(templateStream);
					templateStream.Seek(0, SeekOrigin.Begin);

					using (var readerTemplate = new PdfReader(templateStream))
					{
						using (var pdfStamper = new PdfStamper(readerTemplate, newStream, '1'))
						{
							var over = pdfStamper.GetOverContent(1);
							over.BeginText();
							over.SetFontAndSize(_customFont, 12);
							over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateDate?.ToShortDateString() ?? request.DateChangeStatus?.ToShortDateString() ?? DateTime.Now.ToShortDateString()/*GetDayMonth(request.CertificateDate)*/, 590, 495, 0);
							over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateNumber.FormatEx(), 686, 495, 0);
							over.SetFontAndSize(_customFont, 10);

							float row = 470;

							var requestChildren = request.Child.Where(c => !c.IsDeleted).ToList();

							bool isFirstString = true;
							foreach (var child in requestChildren)
							{

								over.SetFontAndSize(_customFont, 12);
								over.ShowTextAligned(Element.ALIGN_LEFT,
									$"{child.LastName} {child.FirstName} {child.MiddleName}", 180, row, 0);

								WriteByTable(over, _FONT_12, 500, row, 230,
									$"{child.DateOfBirth.FormatEx()}, {child.DocumentSeria} {child.DocumentNumber}",
									Element.ALIGN_CENTER);
								if (isFirstString)
								{
									row = row - 24;
									isFirstString = false;
								}
								else
								{
									row = row - 20.3f;
								}


							}
							over.SetFontAndSize(_customFont, 12);

							WriteByTable(over, _FONT_12, 290, 145, 500, $"{request.RequestNumber}, {applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DocumentType.Name}, {applicant.DocumentSeria} {applicant.DocumentNumber}".Trim());
							over.EndText();

							over = pdfStamper.GetOverContent(2);
							over.BeginText();

							over.SetFontAndSize(_customFont, 12);
							over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateDate?.ToShortDateString() ?? request.DateChangeStatus?.ToShortDateString() ?? DateTime.Now.ToShortDateString()/*GetDayMonth(request.CertificateDate)*/, 590, 495, 0);
							over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateNumber.FormatEx(), 686, 495, 0);
							over.SetFontAndSize(_customFont, 10);

							row = 456;
							isFirstString = true;
							var attendants = request.Attendant.Where(c => !c.IsDeleted).ToList();

							if (request.Applicant.IsAccomp)
							{
								attendants.Add(request.Applicant);
							}

							foreach (var attendant in attendants)
							{

								over.SetFontAndSize(_customFont, 12);
								over.ShowTextAligned(Element.ALIGN_LEFT,
									$"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}", 200, row, 0);

								WriteByTable(over, _FONT_12, 500, row, 230,
									$"{attendant.DateOfBirth.FormatEx()}, {attendant.DocumentSeria} {attendant.DocumentNumber}",
									Element.ALIGN_CENTER);
								if (isFirstString)
								{
									row = row - 27;
									isFirstString = false;
								}
								else
								{
									row = row - 20.3f;
								}

							}

							over.EndText();
						}
					}
				}
			}
			return "Сертификат.pdf";
		}


		/// <summary>
		/// сертификат на отдых 1 ребенок
		/// </summary>
		private string CertificateOnRestSingle(Request request, Stream newStream)
		{
			var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

			var assembly = Assembly.Load("RestChild.Templates");

			var requestChildren = request.Child.Where(c => !c.IsDeleted).ToList();

			using (var stream = assembly.GetManifestResourceStream("RestChild.Templates.single2021.pdf"))
			{
				using (var reader = new PdfReader(stream))
				{
					using (var pdfStamper = new PdfStamper(reader, newStream, '1'))
					{
						var over = pdfStamper.GetOverContent(1);

						var child = requestChildren.FirstOrDefault() ?? new Child
						{ DocumentType = new DocumentType { Name = string.Empty } };

						if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
						{
							child = new Child
							{
								DocumentType =
									request.Applicant?.DocumentType ?? new DocumentType { Name = string.Empty },
								LastName = request.Applicant?.LastName,
								FirstName = request.Applicant?.FirstName,
								MiddleName = request.Applicant?.MiddleName,
								DocumentSeria = request.Applicant?.DocumentSeria,
								DocumentNumber = request.Applicant?.DocumentNumber,
								DateOfBirth = request.Applicant?.DateOfBirth
							};
						}

						var y_delta = 38;

						over.BeginText();
						over.SetFontAndSize(_customFont, 14);
						over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateDate?.ToLongDateString() ?? request.DateChangeStatus?.ToLongDateString() ?? DateTime.Now.ToLongDateString(), 275, 404 + y_delta, 0);
						over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateNumber.FormatEx(), 470, 404 + y_delta, 0);

						var text = $"{child.LastName} {child.FirstName} {child.MiddleName}".Trim();
						WriteByTable(over, text.Length > 108 ? _FONT_10 : _FONT_12, 280, 365 + y_delta, 500, text);
						over.ShowTextAligned(Element.ALIGN_LEFT, $"{child.DateOfBirth.FormatEx()}, {child.DocumentType.Name}, {child.DocumentSeria} {child.DocumentNumber}", 100, 334 + y_delta, 0);

						if (request.TypeOfRestId != (long)TypeOfRestEnum.YouthRestOrphanCamps && (request.TypeOfRest?.NeedAccomodation ?? false))
						{
							var attendant =
								(applicant.IsAccomp && !applicant.IsDeleted
									? applicant
									: request.Attendant.FirstOrDefault(a => a.IsAccomp && !a.IsDeleted)) ??
								new Applicant { DocumentType = new DocumentType { Name = string.Empty } };
							var text_name = $"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}".Trim();
							var text_docs = $"{attendant.DateOfBirth.FormatEx()}, {attendant.DocumentType.Name}, {attendant.DocumentSeria}, {attendant.DocumentNumber}";

							WriteByTable(over, text.Length > 108 ? _FONT_10 : _FONT_12, 400, 295 + y_delta, 500, text_name);
							WriteByTable(over, text.Length > 108 ? _FONT_10 : _FONT_12, 100, 258 + y_delta, 500, text_docs);
						}

						over.ShowTextAligned(Element.ALIGN_LEFT, request.NullSafe(r => r.Tour.DateIncome.Value.ToLongDateString()) ?? string.Empty, 180, 229 + y_delta, 0);

						over.ShowTextAligned(Element.ALIGN_LEFT, request.NullSafe(r => r.Tour.DateOutcome.Value.ToLongDateString()) ?? string.Empty, 300, 229 + y_delta, 0);

						over.ShowTextAligned(Element.ALIGN_LEFT, request.NullSafe(r => r.TransferTo.Name) ?? string.Empty, 300, 155 + y_delta, 0);

						over.ShowTextAligned(Element.ALIGN_LEFT, request.NullSafe(r => r.TransferFrom.Name) ?? string.Empty, 300, 131 + y_delta, 0);

						if (applicant.DocumentType != null)
						{
							text = $"{request.RequestNumber}, {$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}".Trim()}, {applicant.DocumentType.Name}, {applicant.DocumentSeria} {applicant.DocumentNumber}";
							WriteByTable(over, text.Length > 79 ? _FONT_10 : _FONT_12, 260, 106 + y_delta, 550, text);
						}

						over.EndText();

						if (request.Tour != null)
						{
							var headerTable = new PdfPTable(1) { TotalWidth = 600 };
							var p = new Phrase(new Chunk(request.NullSafe(r => r.Tour.Hotels.Name),
								new Font(_customFont, 12)));
							var company = new PdfPCell(p);
							company.SetLeading(11, 0);
							company.HorizontalAlignment = Element.ALIGN_LEFT;
							company.VerticalAlignment = Element.ALIGN_BOTTOM;
							company.BorderWidth = 0;
							headerTable.AddCell(company);
							headerTable.WriteSelectedRows(0, 1, 0, 1, 190, 190 + headerTable.TotalHeight + y_delta, over);
						}
					}
				}
			}
			return "Путёвка.pdf";
		}

		/// <summary>
		/// сертификат на отдых множественный
		/// </summary>
		private string CertificateOnRestMulti(Request request, Stream newStream)
		{

			var assembly = Assembly.Load("RestChild.Templates");

			using (var stream = assembly.GetManifestResourceStream("RestChild.Templates.multi2021.pdf"))
			{
				using (var reader = new PdfReader(stream))
				{
					using (var pdfStamper = new PdfStamper(reader, newStream, '1'))
					{
						var over = pdfStamper.GetOverContent(1);

						var applicant = request.Applicant ?? new Applicant
						{ DocumentType = new DocumentType { Name = string.Empty } };

						var headerTextSize = 12;

						over.BeginText();
						over.SetFontAndSize(_customFont, headerTextSize);
						over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateDate?.ToShortDateString() ?? request.DateChangeStatus?.ToShortDateString() ?? DateTime.Now.ToShortDateString(), 565, 493, 0);
						over.SetFontAndSize(_customFont, headerTextSize);
						over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateNumber.FormatEx(), 694, 493, 0);


						float row = 466;
						bool isFirstString = true;
						var requestChildren = request.Child.Where(c => !c.IsDeleted).ToList();
						foreach (var child in requestChildren)
						{

							over.SetFontAndSize(_customFont, 12);
							over.ShowTextAligned(Element.ALIGN_LEFT,
								$"{child.LastName} {child.FirstName} {child.MiddleName}", 180, row, 0);

							WriteByTable(over, _FONT_12, 500, row, 230,
								$"{child.DateOfBirth.FormatEx()}, {child.DocumentSeria} {child.DocumentNumber}",
								Element.ALIGN_CENTER);
							if (isFirstString)
							{
								row = row - 24;
								isFirstString = false;
							}
							else
							{
								row = row - 20.3f;
							}
						}

						over.SetFontAndSize(_customFont, 12);
						over.ShowTextAligned(Element.ALIGN_LEFT, request.Tour?.DateIncome.Value.ToShortDateString(), 165, 140, 0);

						over.ShowTextAligned(Element.ALIGN_LEFT, request.Tour?.DateOutcome.Value.ToShortDateString(), 270, 140, 0);


						var text = request.Tour?.Hotels?.Name ?? string.Empty;
						WriteByTable(over, _FONT_12, 480, 140, 500, text);

						over.ShowTextAligned(Element.ALIGN_LEFT, request.TransferTo?.Name ?? string.Empty, 285, 100, 0);

						over.ShowTextAligned(Element.ALIGN_LEFT, request.TransferFrom?.Name ?? string.Empty, 285, 75, 0);

						text =
							$"{request.RequestNumber}, {$"123{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}".Trim()}, {applicant.DocumentType?.Name} {applicant.DocumentSeria} {applicant.DocumentNumber}";
						WriteByTable(over, text.Length > 79 ? _FONT_10 : _FONT_12, 220, 47, 550, text, Element.ALIGN_CENTER);

						over.EndText();

						// ---------------------------------------------------
						// вторая страница
						over = pdfStamper.GetOverContent(2);
						over.BeginText();
						over.SetFontAndSize(_customFont, headerTextSize);
						over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateDate?.ToShortDateString() ?? request.DateChangeStatus?.ToShortDateString() ?? DateTime.Now.ToShortDateString(), 570, 490, 0);
						over.SetFontAndSize(_customFont, headerTextSize);
						over.ShowTextAligned(Element.ALIGN_LEFT, request.CertificateNumber.FormatEx(), 692, 490, 0);

						row = 452;
						isFirstString = true;
						var attendants = request.Attendant.Where(c => !c.IsDeleted).ToList();
						if (request.Applicant.IsAccomp)
						{
							attendants.Add(request.Applicant);
						}

						foreach (var attendant in attendants)
						{

							over.SetFontAndSize(_customFont, 12);
							over.ShowTextAligned(Element.ALIGN_LEFT,
								$"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}", 200, row, 0);

							WriteByTable(over, _FONT_12, 500, row, 230,
								$"{attendant.DateOfBirth.FormatEx()}, {attendant.DocumentSeria} {attendant.DocumentNumber}",
								Element.ALIGN_CENTER);
							if (isFirstString)
							{
								row = row - 27;
								isFirstString = false;
							}
							else
							{
								row = row - 20.3f;
							}

						}

						over.EndText();
					}
				}
			}
			return "Путёвка.pdf";
		}

		/// <summary>
		/// Подпись сведений о заявлении
		/// </summary>
		private void SignCertBlock(Document document, Request request)
		{
			var number = "        1090-" + request.RequestNumber.Substring(20, 10) + "         ";
			var date = "        " + DateTime.Now.FormatEx("dd.MM.yyyy") + "        ";
			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = "22" });

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));

			var captionRunProperties = new RunProperties().SetFont().SetFontSizeSupperscript();



			var table = new Table();
			table.AppendChild(tblProp.CloneNode(true));

			//отступ - нулевая строка
			//PdfAddParagraph(document, 0, 20, Element.ALIGN_CENTER, 0);

			PdfAddParagraph(document, 0, 0, Element.ALIGN_JUSTIFIED, 0,
							//    new Chunk("Исполнитель: ", _MAIN_TXT_FONT),
							new Chunk("                                          ", _MAIN_TXT_FONT),
							new Chunk(date, _SMALL_TXT_FONT_UNDERLINED),
							new Chunk("                        ", _MAIN_TXT_FONT),
							new Chunk(number, _SMALL_TXT_FONT_UNDERLINED));
			PdfAddParagraph(document, 0, 0, Element.ALIGN_JUSTIFIED, 0,
							//   new Chunk("                  ", SmallText),
							new Chunk("                                                         ", _SMALL_TXT_FONT),
							new Chunk("(дата регистрации заявления)", _SMALL_TXT_FONT),
							new Chunk("                             ", _SMALL_TXT_FONT),
							new Chunk("(регистрационный номер заявления)", _SMALL_TXT_FONT));


		}

		/// <summary>
		/// Подпись сведений о заявлении
		/// </summary>
		private void AddTableDocsList(Document document, List<string> docs)
		{
			int tablecolls = 9;

			var table = new PdfPTable(tablecolls);
			table.WidthPercentage = 100;
			foreach (string doc in docs)
			{
				table.AddCell("");
				PdfPCell cell = new PdfPCell(new Phrase(10, doc, new Font(_BASE_ITALIC_FONT, 10)));
				cell.Colspan = tablecolls - 1;
				cell.HorizontalAlignment = Element.ALIGN_LEFT;
				cell.VerticalAlignment = Element.ALIGN_CENTER;
				cell.Padding = 2;
				cell.PaddingBottom = 7;
				table.AddCell(cell);
			}

			PdfAddParagraph(document, 0, 10, Element.ALIGN_CENTER, 0);
			document.Add(table);

		}
		#region Others

		private void GetPdfTable(List<Request> requests, Document doc, IEnumerable<int> years)
		{
			PdfPTable pdfTable = new PdfPTable(3);
			pdfTable.SetWidthPercentage(new float[3] { 100, 220, 280 }, PageSize.A4);
			pdfTable.NormalizeHeadersFooters();
			pdfTable.SplitLate = false;
			//pdfTable.DefaultCell.Padding = 3;
			//pdfTable.WidthPercentage = 30;
			//pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
			pdfTable.DefaultCell.BorderWidth = 0.5f;

			PdfPCell pdfPCell; // Создаем ячейку

			pdfPCell = new PdfPCell(new Phrase("Год кампании", _HEADER_FONT));
			pdfPCell.BorderColor = new BaseColor(0, 0, 0); // Цвет границы ячейки, по умолчанию черный
			pdfPCell.BackgroundColor = new BaseColor(255, 255, 255); // Цвет фона ячейки
			pdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
			pdfTable.AddCell(pdfPCell);

			pdfPCell = new PdfPCell(new Phrase("Вид услуги (путевка, сертификат, компенсация)", _HEADER_FONT));
			pdfPCell.BorderColor = new BaseColor(0, 0, 0); // Цвет границы ячейки, по умолчанию черный
			pdfPCell.BackgroundColor = new BaseColor(255, 255, 255); // Цвет фона ячейки
			pdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
			pdfTable.AddCell(pdfPCell);

			pdfPCell = new PdfPCell(new Phrase("Организация отдыха и оздоровления (в случае предоставления путевки для отдыха и оздоровления), даты заезда", _HEADER_FONT));
			pdfPCell.BorderColor = new BaseColor(0, 0, 0); // Цвет границы ячейки, по умолчанию черный
			pdfPCell.BackgroundColor = new BaseColor(255, 255, 255); // Цвет фона ячейки
			pdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
			pdfTable.AddCell(pdfPCell);

			var moneyTypes = new[]
{
				(long?) TypeOfRestEnum.MoneyOn18, (long?) TypeOfRestEnum.MoneyOn3To7,
				(long?) TypeOfRestEnum.MoneyOn7To15, (long?) TypeOfRestEnum.MoneyOnInvalidOn4To17,
				(long?) TypeOfRestEnum.MoneyOnInvalidAndEscort4To17, (long?) TypeOfRestEnum.MoneyOnLimitationAndEscort4To17
			};

			foreach (int year in years)
			{
				Request request = requests.FirstOrDefault(req => req.YearOfRest.Year == year);
				if (!request.IsNullOrEmpty())
				{
					pdfPCell = new PdfPCell(new Phrase(year.ToString(), _MAIN_TXT_FONT));
					pdfPCell.BorderColor = new BaseColor(0, 0, 0);
					pdfPCell.BackgroundColor = new BaseColor(255, 255, 255);
					pdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
					pdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
					pdfTable.AddCell(pdfPCell);

					pdfPCell = new PdfPCell(new Phrase(request?.TypeOfRest?.Name, _MAIN_TXT_FONT));
					pdfPCell.BorderColor = new BaseColor(0, 0, 0);
					pdfPCell.BackgroundColor = new BaseColor(255, 255, 255);
					pdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
					pdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
					pdfTable.AddCell(pdfPCell);

					pdfPCell = new PdfPCell(new Phrase(request == null ? "" :
								(request.TourId.HasValue && !request.Tour.IsNullOrEmpty())
								? $"{request.Tour.Hotels?.Name}, c {request.Tour.DateIncome.FormatEx(string.Empty)} по {request.Tour.DateOutcome.FormatEx(string.Empty)}"
								: !request.TourId.HasValue && (request.TypeOfRest.Id == (long?)TypeOfRestEnum.Money || request.TypeOfRest.ParentId == (long?)TypeOfRestEnum.Money || request.TypeOfRest.Parent.ParentId == (long?)TypeOfRestEnum.Money) && !request.Certificates.IsNullOrEmpty() && !request.Certificates.Where(c => c.Request.YearOfRestId == request.YearOfRestId).FirstOrDefault().IsNullOrEmpty()
								? $"{request.Certificates.Where(c => c.Request.YearOfRestId == request.YearOfRestId).FirstOrDefault().Place}, c {request.Certificates.Where(c => c.Request.YearOfRestId == request.YearOfRestId).FirstOrDefault().RestDateFrom.FormatEx(string.Empty)} по {request.Certificates.Where(c => c.Request.YearOfRestId == request.YearOfRestId).FirstOrDefault().RestDateTo.FormatEx(string.Empty)}"
								: (request.RequestOnMoney && !moneyTypes.Contains(request.TypeOfRestId)
								? "Осуществлен выбор сертификата на втором этапе заявочной кампании" : ""), _MAIN_TXT_FONT));
					pdfPCell.BorderColor = new BaseColor(0, 0, 0);
					pdfPCell.BackgroundColor = new BaseColor(255, 255, 255);
					pdfPCell.HorizontalAlignment = Element.ALIGN_CENTER;
					pdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
					pdfTable.AddCell(pdfPCell);
				}

			}

			doc.Add(pdfTable);
		}

		/// <summary>
		///     Шапка документа
		/// </summary>
		private void PdfAddHeader(Document document)
		{
			PdfAddParagraph(document, "ДЕПАРТАМЕНТ КУЛЬТУРЫ ГОРОДА МОСКВЫ", 0, 1, Element.ALIGN_CENTER, 0, _HEADER_FONT);
			PdfAddParagraph(document, "Государственное автономное учреждение культуры города Москвы", 0, 1,
				Element.ALIGN_CENTER, 0, _HEADER_FONT);
			PdfAddParagraph(document, "\"Московское агентство организации отдыха и туризма\"", 0, 1,
				Element.ALIGN_CENTER, 0, _HEADER_FONT);
			PdfAddParagraph(document, "(ГАУК \"МОСГОРТУР\")", 0, 1, Element.ALIGN_CENTER, 0, _HEADER_FONT);
			var p = new Paragraph(new Chunk(new LineSeparator(0.0F, 100.0F, BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
			document.Add(p);
		}

		/// <summary>
		///     Добавить параграф
		/// </summary>
		private void PdfAddParagraph(Document document, string text, float spacingBefore, float spacingAfter,
			int alignment, float firstLineIndent, Font font)
		{
			var p = new Paragraph(text, font)
			{
				SpacingBefore = spacingBefore,
				SpacingAfter = spacingAfter,
				Alignment = alignment,
				FirstLineIndent = firstLineIndent
			};
			document.Add(p);
		}

		/// <summary>
		///     Добавить параграф
		/// </summary>
		private static void PdfAddParagraph(Document document, float spacingBefore, float spacingAfter, int alignment,
			float firstLineIndent, params Chunk[] chunks)
		{
			var phrase = new Phrase();
			phrase.AddRange(chunks);

			var p = new Paragraph(phrase)
			{
				SpacingBefore = spacingBefore,
				SpacingAfter = spacingAfter,
				Alignment = alignment,
				FirstLineIndent = firstLineIndent
			};
			document.Add(p);
		}

		private void WriteByTable(PdfContentByte over, Font font, float posX, float posY, float width,
			string value, int horizontalAlignment = Element.ALIGN_LEFT)
		{
			over.EndText();
			var table = new PdfPTable(1) { TotalWidth = width };
			var p = new Phrase(new Chunk(value, font));
			var company = new PdfPCell(p)
			{
				HorizontalAlignment = horizontalAlignment,
				VerticalAlignment = Element.ALIGN_TOP,
				BorderWidth = 0
			};

			table.AddCell(company);
			table.WriteSelectedRows(0, 1, 0, 1, posX, posY + table.TotalHeight, over);
			over.BeginText();
		}
		#endregion
		#endregion
	}
}

//DOC генератор
namespace RestChild.DocumentGeneration.DocumentProc.Doc
{
	public class DocProc : BaseDocumentProcessor
	{
		#region ConstSettings

		//размеры шрифтов
		private const string _SIZE_16 = "16";
		private const string _SIZE_18 = "18";
		private const string _SIZE_20 = "20";
		private const string _SIZE_22 = "22";
		private const string _SIZE_24 = "24";
		private const string _SIZE_28 = "28";

		//отступы
		private const string _SPACING_BETWEEN_LINES_AFTER = "20";
		private const string _SPACING_BETWEEN_LINES_LINE = "240";

		//константы для подписи в подвале документа
		private const int _COLUMN_1 = 2731; // исполнитель/ответственный
		private const int _COLUMN_2 = 3931; // ФИО
		private const int _COLUMN_3 = 55; // разделитель
		private const int _COLUMN_4 = _COLUMN_1; // Подпись
		private const int _COLUMN_5 = _COLUMN_3; // второй разделитель
		private const int _COLUMN_6 = 1804; // Дата

		//отступ первой строки
		private const int _FIRST_LINE_INDENTATION_600 = 600;

		private const string _WORDPROCESSINGML_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
		#endregion

		public DocProc() : base(DocTypeEnum.Doc)
		{

		}

		#region Override

		public override IDocument NotificationDeclineNotAccepted(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));



					//    var elems = new List<OpenXmlElement>();
					//    elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold());
					//    elems.Add(new Text("Уведомление о замене сопровождающего лица"));

					//    doc.AppendChild(new Paragraph(
					//        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					//        new Run(elems)
					//    ));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Left },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					//    var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					//    var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					//    titleRequestRunPropertiesBold.AppendChild(new Bold());

					//    var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					//    titleRequestRunPropertiesItalic.AppendChild(new Italic());

					//    var titleRequestRunPropertiesUnderline = titleRequestRunProperties.CloneNode(true);
					//    titleRequestRunPropertiesUnderline.AppendChild(new Indentation());

					//    var documentRunPropertiesItalic = new RunProperties().SetFont().SetFontSize(_SIZE_16);
					//    documentRunPropertiesItalic.AppendChild(new Italic());

					//    var applicant = request.Applicant ??
					//                    new Applicant { DocumentType = new DocumentType { Name = string.Empty } };
					//    Applicant oldAttenadant = null;
					//    Applicant attendant = request.Attendant.FirstOrDefault(att => att.Id == attendantId);
					//    if (request.Applicant.Id == oldAttendantId)
					//        oldAttenadant = applicant;
					//    else oldAttenadant = request.Attendant.FirstOrDefault(att => att.Id == oldAttendantId);

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text(
					//                    $"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text(applicant.Phone + ", " + applicant.Email))));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Left },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Номер путевки: ") { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text(request.RequestNumber.FormatEx()))));


					//    if (!request.Tour.IsNullOrEmpty())
					//        doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Left },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Организация отдыха и оздоровления: ") { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text(request.Tour.Hotels.NameOrganization.FormatEx()))));

					//    if (!request.TimeOfRest.IsNullOrEmpty())
					//        doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Left },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Время отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text(request.TimeOfRest.Name.FormatEx()))));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Left },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Дата замены сопровождающего лица: ")
					//                { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),

					//                new Text(System.DateTime.Now.FormatEx("dd.MM.yyyy")))));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Данные сопровождающего лица, указанного при подаче заявления о предоставлении услуг отдыха и оздоровления: ") { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text(
					//                    $"{oldAttenadant.LastName} {oldAttenadant.FirstName} {oldAttenadant.MiddleName}, {oldAttenadant.DateOfBirth.FormatExGR(string.Empty)}"))));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Данные предлагаемого к замене сопровождающего лица, указанного в заявлении о замене сопровождающего лица: ") { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text(
					//                    $"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}, {attendant.DateOfBirth.FormatExGR(string.Empty)}"))));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Основание приема заявления: ")
					//                { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\"."))));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//                new SpacingBetweenLines { After = _SIZE_20 }),
					//            new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//                new Text("Результат рассмотрения: ")
					//                { Space = SpaceProcessingModeValues.Preserve }),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text("Ваше заявление о замене сопровождающего лица в предоставленной путевке для отдыха и оздоровления рассмотрено положительно. Путевка для отдыха и оздоровления с измененными сведениями о сопровождающем лице прилагается."))
					//            ));

					//    doc.AppendChild(new Paragraph(
					//new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//new SpacingBetweenLines { After = _SIZE_20 })));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//                new SpacingBetweenLines { After = _SIZE_20 },
					//                new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
					//            new Run(titleRequestRunProperties.CloneNode(true),
					//                new Text("Путевка для отдыха и оздоровления с измененными сведениями о сопровождающем лице прилагается.")
					//                { Space = SpaceProcessingModeValues.Preserve })));

					//    doc.AppendChild(new Paragraph(
					//new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//new SpacingBetweenLines { After = _SIZE_20 })));

					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//                new SpacingBetweenLines { After = _SIZE_20 },
					//                new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
					//            new Run(titleRequestRunProperties.CloneNode(true),
					//                new Text("Приложение: путевка для отдыха и оздоровления.")
					//                { Space = SpaceProcessingModeValues.Preserve })));

					//    SignBlockNotification2022(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о непринятии отказа от осуществления отдыха" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		public override IDocument NotificationDeclineAccepted(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var isCert = request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn3To7 ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn7To15 ||
								 //request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn18 || //TODO раскоментировать строку во время альтернативной компании(требуется уточнение)
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalid || почему-то были добавлены эти два типа отдыха, хотя они ни разу не сертификаты
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalidOrphanComplex ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidOn4To17 ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidAndEscort4To17 ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnLimitationAndEscort4To17 ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnBenefits
								 ||
								 request.TypeOfRest.ParentId == (long)TypeOfRestEnum.Money;

					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					if (isCert)
					{
						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text(
									"Уведомление о сформированном отказе от использования сертификата на отдых и оздоровление"))));

						//doc.AppendChild(
						//    new Paragraph(
						//        new ParagraphProperties(new Justification {Val = JustificationValues.Left},
						//            new SpacingBetweenLines {After = _SIZE_20}),
						//        new Run(titleRequestRunPropertiesBold.CloneNode(true),
						//            new Text(
						//                    "Сведения заявления об отказе от использования сертификата на отдых и оздоровление:")
						//                {Space = SpaceProcessingModeValues.Preserve})));
					}
					else
					{
						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text("Уведомление о сформированном отказе от осуществления отдыха"),
								new Break(),
								new Text(
									"и оздоровления на основании предоставленной бесплатной путевки"),
								new Break(),
								new Text(
									"для отдыха и оздоровления"),
								new Break())));
						//doc.AppendChild(
						//    new Paragraph(
						//        new ParagraphProperties(new Justification {Val = JustificationValues.Left},
						//            new SpacingBetweenLines {After = _SIZE_20}),
						//        new Run(titleRequestRunPropertiesBold.CloneNode(true),
						//            new Text("Сведения заявления об отказе от осуществления отдыха и оздоровления:")
						//                {Space = SpaceProcessingModeValues.Preserve})));
					}

					//CertHandInput(doc);


					if (isCert)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер сертификата: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.CertificateNumber.FormatEx()))));
					}
					else
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер путевки: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.CertificateNumber.FormatEx()))));
					}
					if (!request.Tour.IsNullOrEmpty())
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Организация отдыха и оздоровления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.Tour.Hotels.NameOrganization ?? request.Tour.Hotels.Name))));

					if (!request.Tour.IsNullOrEmpty())
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Время отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.Tour.DateIncome.FormatEx("dd.MM.yyyy") + " - " + request.Tour.DateOutcome.FormatEx("dd.MM.yyyy")))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					if (request.Child != null)
					{
						var first = true;
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							if (isCert)
							{
								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Данные лиц, указанных в сертификате: ")
											{ Space = SpaceProcessingModeValues.Preserve },
											new Break()),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text(
												$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));
							}
							else
							{
								var req = child.Request;
								var attendants = new List<Applicant>();
								if (req.Applicant?.IsAccomp ?? false)
								{
									attendants.Add(req.Applicant);
								}
								else
								{
									attendants = req.Attendant.ToList();
								}

								var personsRun = new Run(titleRequestRunPropertiesItalic.CloneNode(true));

								if (first)
								{
									foreach (var at in attendants)
									{
										personsRun.AppendChild(new Text(
											$"{at.LastName} {at.FirstName} {at.MiddleName}, {at.DateOfBirth.FormatExGR(string.Empty)}"));
										personsRun.AppendChild(new Break());
									}
								}

								personsRun.AppendChild(new Text(
									$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"));

								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text(first
												? "Данные лиц, указанных в путевке: "
												: string.Empty)
											{ Space = SpaceProcessingModeValues.Preserve }, first ? new Break() : null),
										personsRun
										));
							}

							if (first)
							{
								first = false;
							}
						}
					}


					if (isCert)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения заявления: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(
										"Заявление об отказе от предоставленного сертификата на отдых и оздоровление удовлетворено. Сертификат на отдых и оздоровление отозван по инициативе заявителя."))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(
										//"постановление Правительства Москвы от 22 февраля 2017 г. № 56-ПП "
										"пункт 8(1).10 Порядка организации отдыха и оздоровления детей, находящихся"),
									new Break(),
									new Text(
										"в трудной жизненной ситуации, утверждённого постановлением Правительства Москвы "),
									new Break(),
									new Text(
										"от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся "),
									new Break(),
									new Text(
										"в трудной жизненной ситуации\"."))));
					}
					else
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения заявления: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(
										"Причина отказа от предоставленной путевки для отдыха и оздоровления признана уважительной. Отказ от осуществления отдыха и оздоровления на основании предоставленной бесплатной путевки для отдыха и оздоровления сформирован."))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(
										"пункты 10.1 и 10.2 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы"),
									new Break(),
									new Text(
										"от 22 февраля 2017 г. № 56 - ПП \"Об организации отдыха и оздоровления детей, находящихся"),
									new Break(),
									new Text(
										"в трудной жизненной ситуации\"."))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					SignBlockNotification2022(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в установленный срок" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		#region NotificationRegistration
		protected override IDocument NotificationBasicRegistration(Request request, Account account)
		{
			var youth = request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps;

			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));


					if (!youth)
					{
						var elems = new List<OpenXmlElement>();
						elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold());
						elems.Add(new Text("Уведомление о регистрации заявления о предоставлении услуг отдыха"));
						elems.Add(new Break());
						elems.Add(new Text("и оздоровления, поданного с помощью работника"));
						elems.Add(new Break());
						elems.Add(new Text("ГАУК \"МОСГОРТУР\""));

						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(elems)
						));
					}
					else
					{
						var elems = new List<OpenXmlElement>();
						elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold());
						elems.Add(new Text("Уведомление о регистрации заявления о предоставлении услуг отдыха"));
						elems.Add(new Break());
						elems.Add(new Text("и оздоровления лицу из числа детей-сирот и детей, оставшихся"));
						elems.Add(new Break());
						elems.Add(new Text("без попечения родителей, поданного с помощью работника"));
						elems.Add(new Break());
						elems.Add(new Text("ГАУК \"МОСГОРТУР\""));

						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(elems)
						));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var documentRunPropertiesItalic = new RunProperties().SetFont().SetFontSize(_SIZE_16);
					documentRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(youth
									? "бесплатная путевка для отдыха и оздоровления"
									: request.TypeOfRest?.Name))));

					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Данные ребёнка: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text($"{child.BenefitType?.Name}"))));
					}


					if (youth)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text(
											"Данные лица из числа детей-сирот и детей, оставшихся без попечения родителей:")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Break(),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Перечень предоставленных заявителем документов: ")
								{ Space = SpaceProcessingModeValues.Preserve })));

					var docs = youth
						? new List<string>
							{
								"Документ, удостоверяющий личность лица из числа детей-сирот и детей, оставшихся без попечения родителей",
								"Документ, удостоверяющий личность доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления доверенным лицом на совершение действий в период проведения заявочной кампании)",
								"Страховой номер индивидуального лицевого счета (СНИЛС) в системе индивидуального (персонифицированного) учета лица из числа детей-сирот и детей, оставшихся без попечения родителей",
								"Страховой номер индивидуального лицевого счета (СНИЛС) в системе индивидуального (персонифицированного) учета доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления доверенным лицом на совершение действий в период проведения заявочной кампании)",
								"Документ, подтверждающий сведения об отнесении лица к категории лиц из числа детей-сирот и детей, оставшихся без попечения родителей",
								"Документ, подтверждающий сведения о месте жительства лица из числа детей-сирот и детей, оставшихся без попечения родителей, в городе Москве",
								"Документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления доверенным лицом на совершение действий в период проведения заявочной кампании)"
							}
						: new List<string>
							{
								"Документ, удостоверяющий личность родителя ребёнка или иного законного представителя – опекуна, попечителя, приемного родителя, патронатного воспитателя ребенка",
								"Документ, удостоверяющий личность доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления доверенным лицом на совершение действий в период проведения заявочной кампании)",
								"Страховой номер индивидуального лицевого счета (СНИЛС) в системе индивидуального (персонифицированного) учета родителя или законного представителя ребёнка",
								"Страховой номер индивидуального лицевого счета (СНИЛС) в системе индивидуального (персонифицированного) учета доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления доверенным лицом на совершение действий в период проведения заявочной кампании)",
                                //"$Документ, удостоверяющий личность ребенка:",
                                "Документ, удостоверяющий личность ребёнка",
                                //"#для ребенка в возрасте до 14 лет – свидетельство о рождении ребенка или документ, подтверждающий факт рождения и регистрации ребенка, выданный в установленном порядке (в случае рождения ребенка на территории иностранного государства);",
                                //"#для ребенка, достигшего возраста 14 лет, – паспорт гражданина Российской Федерации, паспорт гражданина иностранного государства (в случае наличия гражданства иностранного государства)",
                                "Документ, подтверждающий сведения об отнесении ребёнка к одной из категорий детей, находящихся в трудной жизненной ситуации и указанных в пунктах 3.1.2-3.1.13 Порядка",
								"Документ, подтверждающий сведения о месте жительства ребёнка, в городе Москве",
								"Страховой номер индивидуального лицевого счета (СНИЛС) в системе индивидуального (персонифицированного) учета ребёнка",
								"Сведения документа, удостоверяющего личность сопровождающего лица (в случае организации совместного выездного отдыха)",
								"Страховой номер индивидуального лицевого счета (СНИЛС) в системе индивидуального (персонифицированного) учета сопровождающего лица (в случае организации совместного выездного отдыха)",
								"Документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления доверенным лицом на совершение действий в период проведения заявочной кампании)",
								"Документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребёнка доверенным лицом для сопровождения во время отдыха и оздоровления)",
								"Документы, подтверждающие полномочия заявителя, сопровождающего лица (в случае организации совместного выездного отдыха) из числа законных представителей – родителей, опекунов, попечителей, приемных родителей, патронатных воспитателей ребёнка"
							};

					AddTableDocsList(doc, docs);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("Ваше заявление о предоставлении услуг отдыха и оздоровления (далее – заявление) ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text(
										"зарегистрировано")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(" и является ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text(
										"неотъемлемой частью")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										" данного уведомления.")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("В случае если в срок, не превышающий 21 рабочий день с даты подачи, по Вашему заявлению ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("не пришло уведомление об отказе в предоставлении услуг отдыха и оздоровления")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text($" по причинам, указанным в пунктах {(youth ? "9.1.2-9.1.4" : "9.1.2-9.1.5, 9.1.8-9.1.11")} Порядка организации отдыха")
								{ Space = SpaceProcessingModeValues.Preserve },
								new Text(" и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее – Порядок), ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Ваше заявление прошло все проверки")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(" и будет рассмотрено на заседании Комиссии ГАУК \"МОСГОРТУР\" по включению сведений в предварительный Реестр получателей услуг отдыха и оздоровления (далее – Комиссия), которое состоится 16 января 2023 г.")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										"По результатам заседания Комиссии на электронную почту, указанную в заявлении, не позднее ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text(
										"6 февраля 2023 г.")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										youth
											? " будет направлено уведомление о необходимости выбора организации отдыха и оздоровления (допуск на второй этап заявочной кампании)."
											: " будет направлено одно из уведомлений: о необходимости выбора организации отдыха и оздоровления (допуск на второй этап заявочной кампании); о предоставлении сертификата на отдых и оздоровление; об отказе в предоставлении услуг отдыха и оздоровления по причине, указанной в пункте 9.1.1 Порядка.")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					if (!youth)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("При рассмотрении заявлений Комиссия руководствуется пунктом 3.9 Порядка, согласно которому в первую очередь рассматриваются заявления, в которых указан хотя бы один ребёнок, которому в период с 2020 по 2022 год не предоставлялись бесплатные путевки для отдыха и оздоровления (далее – бесплатные путевки) либо сертификаты на отдых и оздоровление (далее – сертификаты), а также не выплачивалась компенсация за самостоятельно приобретенную путевку (далее – компенсация).")
									{ Space = SpaceProcessingModeValues.Preserve })
							));
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("Во вторую очередь рассматриваются заявления, где указаны дети, которым в любые два года из последних трех лет (с 2020 по 2022 год) не предоставлялись бесплатные путевки либо сертификаты, а также не выплачивалась компенсация.")
									{ Space = SpaceProcessingModeValues.Preserve })
							));
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("В третью очередь рассматриваются заявления, где указаны дети, которым в любые два года из последних трех лет (с 2020 по 2022 год) предоставлялись бесплатные путевки либо сертификаты, а также выплачивалась компенсация.")
									{ Space = SpaceProcessingModeValues.Preserve })
							));
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("Исключение составляют дети-сироты и дети, оставшиеся без попечения родителей, находящиеся под опекой, попечительством, в том числе в приемной или патронатной семье, и лица из числа детей-сирот и детей, оставшихся без попечения родителей, которым федеральным законодательством гарантировано ежегодное предоставление услуг отдыха и оздоровления.")
									{ Space = SpaceProcessingModeValues.Preserve })
							));


						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("По результатам второго этапа заявочной кампании, проводимого в период с 7 по 21 февраля 2023 г., уведомление о результате рассмотрения заявления (о предоставлении бесплатной путевки для отдыха и оздоровления; об отказе в предоставлении услуг отдыха и оздоровления в связи с неучастием заявителя во втором этапе заявочной кампании; о неучастии заявителя в выборе конкретной организации отдыха и оздоровления; об отказе в предоставлении услуг отдыха и оздоровления) будет направлено на электронную почту, указанную в заявлении, ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("не позднее 22 февраля 2023 г.")
									{ Space = SpaceProcessingModeValues.Preserve })
							));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("Обращаем внимание, что Вы согласились с необходимостью соблюдения: Правил отдыха и оздоровления, Соглашения об осуществлении обязанностей по сопровождению детей во время отдыха и оздоровления родителем или иным законным представителем либо доверенным лицом для сопровождения во время отдыха и оздоровления (для совместного выездного отдыха); а также проинформированы: о перечне противопоказаний к отдыху и оздоровлению, утвержденном Министерством здравоохранения Российской Федерации (для индивидуального выездного отдыха), об отсутствии возможности замены сертификата на отдых и оздоровление на бесплатную путевку для отдыха и оздоровления (в случае выбора сертификата на отдых и оздоровление), об отсутствии возможности замены бесплатной путевки для отдыха и оздоровления на сертификат на отдых и оздоровление (в случае выбора бесплатной путевки для отдыха и оздоровления), об отсутствии возможности выбора сертификата на отдых и оздоровление в случае отказа на втором этапе заявочной кампании от всех предложенных организаций отдыха и оздоровления, о перечислении денежных средств организации, осуществляющей туроператорскую деятельность и (или) оказывающей услуги отдыха и оздоровления, при приобретении у такой организации услуг отдыха и оздоровления или туристского продукта для ребенка, либо ребёнка и лица, его сопровождающего с использованием сертификата на отдых и оздоровление (в случае выбора сертификата на отдых и оздоровление).")
									{ Space = SpaceProcessingModeValues.Preserve })
							));
					}
					else
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("По результатам второго этапа заявочной кампании, проводимого в период с 7 по 21 февраля 2023 г., уведомление о результате рассмотрения заявления (о предоставлении бесплатной путевки для отдыха и оздоровления; об отказе в предоставлении услуг отдыха и оздоровления в связи с неучастием заявителя во втором этапе заявочной кампании; об отказе в предоставлении услуг отдыха и оздоровления) будет направлено на электронную почту, указанную в заявлении, ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("не позднее 22 февраля 2023 г.")
									{ Space = SpaceProcessingModeValues.Preserve })
							));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("Обращаем внимание, что Вы согласились с необходимостью соблюдения Правил отдыха и оздоровления.")
									{ Space = SpaceProcessingModeValues.Preserve })
							));

					}

					SignBlockNotification2020(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о регистрации" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRegistrationDecline(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text(
								"Уведомление об отказе в регистрации заявления о предоставлении услуг отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата обращения: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные обратившегося лица: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Цель обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("получение услуг отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));

					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text(
												"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));
						}
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text(
											"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание отказа в регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									"Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее — Порядок)."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Причина отказа в регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									//"пункт 5.11.1 Порядка: \"Наличие в отношении одного и того же ребёнка, лица из числа детей-сирот и детей, оставшихся без попечения родителей, другого заявления о предоставлении услуг отдыха и оздоровления в текущем календарном году\"."
									request.DeclineReason == null ? "Заявление является повторным. На указанного(ых) в заявлении ребенка(детей) (ФИО) уже подано заявление о предоставлении  услуг отдыха и оздоровления. " : request.DeclineReason.Name
									))));

					SignBlockNotification2020(doc, account, "Исполнитель:");

					mainPart.Document = doc;

				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в регистрации заявления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationCompensationRegistrationOrphans(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление о регистрации заявления на выплату компенсации"),
							new Break(),
							new Text("за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления"),
							new Break(),
							new Text("(дети-сироты и дети, оставшиеся без попечения родителей, находящиеся под опекой, попечительством, в том числе"),
							new Break(),
							new Text("в приемной или патронатной семье)"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));

					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Данные ребёнка: ") { Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Льготная категория ребёнка: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{child.BenefitType?.Name}"))));
						}
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Перечень предоставленных заявителем документов: ")
								{ Space = SpaceProcessingModeValues.Preserve })));

					var docs = new List<string>
					{
						"заявление на выплату компенсации за самостоятельно приобретенную путевку",
						"документ, удостоверяющий личность обратившегося лица",
						"документ, подтверждающий полномочия законного представителя – опекуна, попечителя, приемного родителя, патронатного воспитателя",
						"$документ, удостоверяющий личность ребёнка-сироты, ребенка, оставшегося без попечения родителей: ",
						"#для ребёнка в возрасте до 14 лет – свидетельство о рождении ребенка или документ, подтверждающий факт рождения и регистрации ребенка, выданный в установленном порядке (в случае рождения ребенка на территории иностранного государства)",
						"#для ребёнка, достигшего возраста 14 лет, – паспорт гражданина Российской Федерации, паспорт гражданина иностранного государства (в случае наличия гражданства иностранного государства)",
						"страховой номер обязательного пенсионного страхования (СНИЛС) ребёнка-сироты, ребенка, оставшегося без попечения родителей",
						"страховой номер обязательного пенсионного страхования (СНИЛС) законного представителя ребёнка-сироты, ребенка, оставшегося без попечения родителей",
						"страховой номер обязательного пенсионного страхования (СНИЛС) сопровождающего лица (в случае сопровождения ребёнка-сироты, ребенка, оставшегося без попечения родителей во время отдыха и оздоровления)",
						"документы, подтверждающие отдых и оздоровление ребёнка-сироты, ребенка, оставшегося без попечения родителей, и оплату законным представителем услуг отдыха и оздоровления",
						"сведения о кредитной организации и открытом счете законного представителя в кредитной организации для выплаты компенсации за самостоятельно приобретенную законными представителями путевку для отдыха и оздоровления",
						"доверенность, подтверждающая полномочия доверенного лица законного представителя (в случае подачи заявления на получение компенсации за приобретенную путевку доверенным лицом законного представителя ребёнка)"
					};

					AddTableDocsList(doc, docs);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text(
										"Заявление")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										" на выплату компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления (дети - сироты и дети, оставшиеся без попечения родителей, находящиеся под опекой, попечительством, в том числе в приемной или патронатной семье) ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("является неотъемлемой частью данного уведомления")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(".")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										"О результате рассмотрения Вы будете уведомлены способом и формой информирования, указанной в заявлении.")
								{ Space = SpaceProcessingModeValues.Preserve })));


					SignBlockNotification2019(doc, account, "Подпись заявителя, подтверждающая получение уведомления о регистрации заявления на выплату компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления (дети-сироты и дети, оставшиеся без попечения родителей, находящиеся под опекой, попечительством, в том числе в приемной или патронатной семье):");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о регистрации" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationCompensationRegistrationForPoors(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление о регистрации заявления на выплату компенсации"),
							new Break(),
							new Text("за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления"),
							new Break(),
							new Text("(дети из малообеспеченных семей)"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));

					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Данные ребёнка: ") { Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Льготная категория ребёнка: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{child.BenefitType?.Name}"))));
						}
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Перечень предоставленных заявителем документов: ")
								{ Space = SpaceProcessingModeValues.Preserve })));

					var docs = new List<string>
					{
						"заявление на выплату компенсации за самостоятельно приобретенную путевку",
						"документ, удостоверяющий личность обратившегося лица",
						"документ, подтверждающий полномочия законного представителя – опекуна, попечителя, приемного родителя, патронатного воспитателя (в случае подачи заявления на выплату компенсации за приобретенную путевку обратившимся лицом из числа законных представителей – опекунов, попечителей, приемных родителей, патронатных воспитателей ребёнка)",
						"$документ, удостоверяющий личность ребёнка: ",
						"#для ребёнка в возрасте до 14 лет – свидетельство о рождении ребенка или документ, подтверждающий факт рождения и регистрации ребенка, выданный в установленном порядке (в случае рождения ребенка на территории иностранного государства)",
						"#для ребёнка, достигшего возраста 14 лет, – паспорт гражданина Российской Федерации, паспорт гражданина иностранного государства (в случае наличия гражданства иностранного государства)",
						"страховой номер обязательного пенсионного страхования (СНИЛС) ребёнка",
						"страховой номер обязательного пенсионного страхования (СНИЛС) родителя, иного законного представителя",
						"документы, подтверждающие факт родства родителя с ребёнком (в случае если иные представленные документы не содержат необходимой информации)",
						"документ, содержащий сведения о месте жительства ребёнка в городе Москве (в случае если в документе, удостоверяющем личность ребенка, отсутствуют сведения о его месте жительства в городе Москве)",
						"документы, подтверждающие отдых и оздоровление ребёнка и оплату родителем услуг отдыха и оздоровления",
						"сведения о кредитной организации и открытом счете родителя, иного законного представителя в кредитной организации для выплаты компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления",
						"доверенность, подтверждающая полномочия доверенного лица родителя (в случае подачи заявления на получение компенсации за приобретенную путевку доверенным лицом родителя, иного законного представителя ребёнка)"
					};

					AddTableDocsList(doc, docs);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text(
										"Заявление")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										" на выплату компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления (дети из малообеспеченных семей) ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("является неотъемлемой частью данного уведомления")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(".")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										"О результате рассмотрения Вы будете уведомлены способом и формой информирования, указанной в заявлении.")
								{ Space = SpaceProcessingModeValues.Preserve })));


					SignBlockNotification2019(doc, account, "Подпись заявителя, подтверждающая получение уведомления о регистрации заявления на выплату компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления (дети из малообеспеченных семей):");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о регистрации" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationCompensationRegistration(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление о регистрации заявления на выплату компенсации"),
							new Break(),
							new Text("за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));

					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text(
												"Данные лица из числа детей-сирот, и детей оставшихся без попечения родителей: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Льготная категория: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{child.BenefitType?.Name}"))));
						}
					}

					//Наркоманы и МГТ хотя чтобы было так
					if (request.Child == null || !request.Child.Any())
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text(
											"Данные лица из числа детей-сирот, и детей оставшихся без попечения родителей: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text($"{applicant.BenefitType?.Name}"))));
					}


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Перечень предоставленных заявителем документов: ")
								{ Space = SpaceProcessingModeValues.Preserve })));

					var docs = new List<string>
					{
						"документ, удостоверяющий личность обратившегося лица",
						"страховой номер обязательного пенсионного страхования (СНИЛС) лица из числа детей-сирот и детей, оставшихся без попечения родителей",
						"доверенность, подтверждающая полномочия представителя, уполномоченного лицом из числа детей-сирот и детей, оставшихся без попечения родителей, на подачу заявления на выплату компенсации за самостоятельно приобретенную путевку (в случае подачи заявления на выплату компенсации за самостоятельно приобретенную путевку таким представителем)",
						"страховой номер обязательного пенсионного страхования (СНИЛС) представителя, уполномоченного лицом из числа детей-сирот и детей, оставшихся без попечения родителей, на подачу заявления на выплату компенсации за самостоятельно приобретенную путевку (в случае подачи заявления на выплату компенсации за самостоятельно приобретенную путевку таким представителем)",
						"документы, подтверждающие отдых и оздоровление, и оплату услуг отдыха и оздоровления лицом из числа детей-сирот и детей, оставшихся без попечения родителей",
						"сведения о кредитной организации и открытом лица из числа детей-сирот и детей, оставшихся без попечения родителей для выплаты компенсации за самостоятельно приобретенную лицом их числа детей-сирот и детей, оставшихся без попечения родителей путевку для отдыха и оздоровления"
					};

					AddTableDocsList(doc, docs);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text(
										"Заявление")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										" на выплату компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("является неотъемлемой частью данного уведомления")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(".")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										"О результате рассмотрения Вы будете уведомлены способом и формой информирования, указанной в заявлении.")
								{ Space = SpaceProcessingModeValues.Preserve })));

					SignBlockNotification2019(doc, account, "Подпись заявителя, подтверждающая получение уведомления о регистрации заявления на выплату компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о регистрации" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}
		#endregion

		//NotImplemented
		public override IDocument CertificateForRequestTemporaryFile(Request request, long? sendStatusId = null)
		{
			throw new NotImplementedException();
		}

		#region NotificationRefuse

		protected override IDocument NotificationDeadline(Request request, Account account)
		{

			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var isCert = request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn3To7 ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn7To15 ||
								 //request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn18 || //TODO раскоментировать строку во время альтернативной компании(требуется уточнение)
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalid || почему-то были добавлены эти два типа отдыха, хотя они ни разу не сертификаты
								 // request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalidOrphanComplex ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidOn4To17 ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidAndEscort4To17 ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnLimitationAndEscort4To17 ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnBenefits
								 ||
								 request.TypeOfRest.ParentId == (long)TypeOfRestEnum.Money; ;

					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					if (isCert)
					{
						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text(
									"Уведомление о сформированном отказе от использования сертификата на отдых и оздоровление"))));

						//doc.AppendChild(
						//    new Paragraph(
						//        new ParagraphProperties(new Justification {Val = JustificationValues.Left},
						//            new SpacingBetweenLines {After = _SIZE_20}),
						//        new Run(titleRequestRunPropertiesBold.CloneNode(true),
						//            new Text(
						//                    "Сведения заявления об отказе от использования сертификата на отдых и оздоровление:")
						//                {Space = SpaceProcessingModeValues.Preserve})));
					}
					else
					{
						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text("Уведомление о сформированном отказе от осуществления отдыха"),
								new Break(),
								new Text(
									"и оздоровления на основании предоставленной бесплатной путевки"),
								new Break(),
								new Text(
									"для отдыха и оздоровления"),
								new Break())));
						//doc.AppendChild(
						//    new Paragraph(
						//        new ParagraphProperties(new Justification {Val = JustificationValues.Left},
						//            new SpacingBetweenLines {After = _SIZE_20}),
						//        new Run(titleRequestRunPropertiesBold.CloneNode(true),
						//            new Text("Сведения заявления об отказе от осуществления отдыха и оздоровления:")
						//                {Space = SpaceProcessingModeValues.Preserve})));
					}

					//CertHandInput(doc);


					if (isCert)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер сертификата: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.CertificateNumber.FormatEx()))));
					}
					else
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер путевки: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.CertificateNumber.FormatEx()))));
					}
					if (!request.Tour.IsNullOrEmpty())
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Организация отдыха и оздоровления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.Tour.Hotels.NameOrganization ?? request.Tour.Hotels.Name))));

					if (!request.Tour.IsNullOrEmpty())
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Время отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.Tour.DateIncome.FormatEx("dd.MM.yyyy") + " - " + request.Tour.DateOutcome.FormatEx("dd.MM.yyyy")))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					if (request.Child != null)
					{
						var first = true;
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							if (isCert)
							{
								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Данные лиц, указанных в сертификате: ")
											{ Space = SpaceProcessingModeValues.Preserve },
											new Break()),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text(
												$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));
							}
							else
							{
								var req = child.Request;
								var attendants = new List<Applicant>();
								if (req.Applicant?.IsAccomp ?? false)
								{
									attendants.Add(req.Applicant);
								}
								else
								{
									attendants = req.Attendant.ToList();
								}

								var personsRun = new Run(titleRequestRunPropertiesItalic.CloneNode(true));

								if (first)
								{
									foreach (var at in attendants)
									{
										personsRun.AppendChild(new Text(
											$"{at.LastName} {at.FirstName} {at.MiddleName}, {at.DateOfBirth.FormatExGR(string.Empty)}"));
										personsRun.AppendChild(new Break());
									}
								}

								personsRun.AppendChild(new Text(
									$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"));

								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text(first
												? "Данные лиц, указанных в путевке: "
												: string.Empty)
											{ Space = SpaceProcessingModeValues.Preserve }, first ? new Break() : null),
										personsRun
										));
							}

							if (first)
							{
								first = false;
							}
						}
					}


					if (isCert)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения заявления: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(
										"Заявление об отказе от предоставленного сертификата на отдых и оздоровление удовлетворено. Сертификат на отдых и оздоровление отозван по инициативе заявителя."))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(
										//"постановление Правительства Москвы от 22 февраля 2017 г. № 56-ПП "
										"пункт 8(1).10 Порядка организации отдыха и оздоровления детей, находящихся"),
									new Break(),
									new Text(
										"в трудной жизненной ситуации, утверждённого постановлением Правительства Москвы "),
									new Break(),
									new Text(
										"от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся "),
									new Break(),
									new Text(
										"в трудной жизненной ситуации\"."))));
					}
					else
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения заявления: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(
										"Причина отказа от предоставленной путевки для отдыха и оздоровления признана уважительной. Отказ от осуществления отдыха и оздоровления на основании предоставленной бесплатной путевки для отдыха и оздоровления сформирован."))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(
										"пункты 10.1 и 10.2 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы"),
									new Break(),
									new Text(
										"от 22 февраля 2017 г. № 56 - ПП \"Об организации отдыха и оздоровления детей, находящихся"),
									new Break(),
									new Text(
										"в трудной жизненной ситуации\"."))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					SignBlockNotification2022(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в установленный срок" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse1090(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление о прекращении рассмотрения поданного заявления"),
							new Break(),
							new Text("о предоставлении услуг отдыха и оздоровления")
						)));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);

					var applicant = request.NullSafe(r => r.Applicant) ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text($"{applicant.Phone}, {applicant.Email}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Причина обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"прекращение рассмотрения поданного заявления о предоставлении услуг отдыха и оздоровления по инициативе заявителя."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание обращения: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"пункт 5.13 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\": \"Заявитель вправе отозвать заявление о предоставлении услуг отдыха и оздоровления в срок не позднее дня окончания стадии заявочной кампании по приему заявлений о предоставлении услуг отдыха и оздоровления (до 10 декабря {DateTime.Now.Year} г. включительно)\".")
							)));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Информация о заявлении:") { Space = SpaceProcessingModeValues.Preserve })));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{
									Space = SpaceProcessingModeValues.Preserve
								}),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("отозвано по инициативе заявителя в установленный срок."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата прекращения рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateChangeStatus.FormatEx(string.Empty)))));

					SignBlockNotification(doc, account, $"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о прекращении рассмотрения поданного заявления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse10802(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление об отказе в предоставлении услуг отдыха и оздоровления"),
							new Break(),
							new Text("в связи с представлением документов, не соответствующих требованиям"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));


					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text(
												"Данные ребёнка/лица из числа детей сирот и детей, оставшихся без попечения родителей: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Льготная категория: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{child.BenefitType?.Name}"))));
						}
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("отказ в предоставлении услуг отдыха и оздоровления."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание отказа: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									"Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее – Порядок)."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Причина отказа: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(request.DeclineReason?.Name))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Уважаемый заявитель!") { Space = SpaceProcessingModeValues.Preserve })));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Просим обратить внимание на следующую информацию.")
								{ Space = SpaceProcessingModeValues.Preserve })));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
										"В случае принятия решения о повторной подаче заявления, рекомендуем убедиться в корректности внесенных в заявление сведений.")
								{ Space = SpaceProcessingModeValues.Preserve })));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
										"Обращаем внимание, что сведения, указанные в заявлении, должны полностью соответствовать сведениям, указанным в СНИЛС и документе, удостоверяющем личность заявителя, ребёнка/детей (в том числе следует обратить внимание на соответствие написания букв \"е - ё\").")
								{ Space = SpaceProcessingModeValues.Preserve })));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
										"В случае если сведения о Вас и ребёнке/детях внесены в \"личный кабинет\" Портала Мэра и Правительства Москвы (при подаче заявления данные отразились в полях автоматически), рекомендуем Вам проверить корректность внесенных сведений, а именно: сверить данные, указанные в \"личном кабинете\" Портала Мэра и Правительства Москвы и данные, указанные в СНИЛС и документе, удостоверяющем личность: фамилия, имя, отчество, пол, дата рождения.")
								{ Space = SpaceProcessingModeValues.Preserve })));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
										"В случае если Вами обнаружена ошибка в \"личном кабинете\" Портала Мэра и Правительства Москвы, ее необходимо исправить.")
								{ Space = SpaceProcessingModeValues.Preserve })));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
										"Дополнительно сообщаем, что в случае обнаружения расхождений сведений в СНИЛС и документе, удостоверяющем личность, Вам необходимо актуализировать документы в соответствующих органах, приведя их к единообразию.")
								{ Space = SpaceProcessingModeValues.Preserve })));

					//SignBlockNotification(doc, account, $"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}");
					SignBlockNotification2020(doc, account, "Исполнитель");
					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в предоставлении услуг отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse10805(Request request, Account account)
		{
			IUnitOfWork unitOfWork = new UnitOfWork();
			var lokYear = request.YearOfRest?.Year ?? 2021;
			var yearIds = unitOfWork.GetSet<YearOfRest>().Where(ss => ss.Year < lokYear).OrderByDescending(ss => ss.Year)
														 .Take(3).Select(ss => ss.Id).ToList();
			IEnumerable<int> years = unitOfWork.GetSet<YearOfRest>().Where(x => yearIds.Contains(x.Id)).Select(x => x.Year).OrderBy(x => x).ToList();
			var listTravelersRequest = unitOfWork.GetSet<ListTravelersRequest>();

			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text(
								"Уведомление о результатах рассмотрения заявления о предоставлении услуг отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"бесплатная путевка для отдыха и оздоровления/сертификат на отдых и оздоровление"))));


					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Данные ребёнка: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория ребёнка: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text($"{child.BenefitType?.Name}"))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									$"учитывая получение услуг отдыха и оздоровления в предыдущие годы и исходя из даты и времени подачи заявления, оказание услуг отдыха и оздоровления в 2023 году не представляется возможным."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									"пункты 3.9. и 9.1.1. Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\"."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text(
										"Информация об услугах, оказанных ребёнку/детям в течение последних 3-х лет")
								{ Space = SpaceProcessingModeValues.Preserve })));


					//if ((request.Child?.Any(c => !c.IsDeleted) ?? false) && listTravelersRequest != null &&
					//    listTravelersRequest.Details.Any(ss => ss.Detail != "[]"))
					//{

					//var details = listTravelersRequest?.Details.Where(ss => ss.Detail != "[]")
					//    .Select(ss => ss.Detail).ToList().SelectMany(JsonConvert.DeserializeObject<DetailInfo[]>)
					//    .ToArray();

					//var details = listTravelersRequest?.SelectMany(d => d.Details)
					//                                   .Where(ss => ss.Detail != "[]")
					//                                   .Select(ss => ss.Detail)
					//                                   .SelectMany(JsonConvert.DeserializeObject<DetailInfo[]>).ToList();


					//IEnumerable<int> years = unitOfWork.GetSet<YearOfRest>().Where(x => yearIds.Contains(x.Id)).Select(x => x.Year).OrderBy(x=>x).ToList();

					IEnumerable<Request> requests = new List<Request>();

					foreach (var child in request.Child.Where(c => !c.IsDeleted).ToList())
					{
						//сравнение по снилсам потому что структура БД наркоманская
						var sameChilds = unitOfWork.GetSet<Child>().Where(ch => ch.Snils == child.Snils && ch.IsDeleted != true).ToList();

						var requestIds = sameChilds?.Select(ss => ss.RequestId).Distinct().ToList();

						requestIds.Remove(request.Id);


						requests = unitOfWork.GetSet<Request>().Where(re => requestIds.Any(req => req == re.Id)).ToList();

						var initialRequests = requests.Where(re => re.StatusId == (long)StatusEnum.Reject).ToList();

						var descendantRequests = unitOfWork.GetSet<Request>().Where(ss => ss.ParentRequestId != null && requestIds.Contains((long)ss.ParentRequestId) && years.Contains(ss.YearOfRest.Year) && ss.StatusId == (long)StatusEnum.CertificateIssued).ToList();


						var resultRequests = initialRequests.Join(descendantRequests, r => r.Id, dr => dr.ParentRequestId, (r, dr) => new Request { Id = dr.Id, Tour = dr.Tour, TourId = r.TourId, TypeOfRest = dr.TypeOfRest, RequestOnMoney = r.RequestOnMoney, TypeOfRestId = r.TypeOfRestId, YearOfRest = r.YearOfRest }).ToList();
						resultRequests.AddRange(requests.Where(req => !resultRequests.Any(rr => rr.Id == req.Id) && req.StatusId == (long?)StatusEnum.CertificateIssued));
						doc.AppendChild(
							new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										  new SpacingBetweenLines { After = _SIZE_20 }),
										  new Run(titleRequestRunProperties.CloneNode(true),
												  new Text($"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}, {child.Snils}"))));

						AddTableChildTours(doc, resultRequests, years);

					}

					SignWorkerBlock(doc, account, "Исполнитель:");
					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о результатах рассмотрения заявления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse108013(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление об отказе в предоставлении услуг отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));


					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text(
												"Данные ребенка (детей) /лица из числа детей-сирот и детей, оставшихся без попечения родителей: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Льготная категория: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{child.BenefitType?.Name}"))));
						}
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text(
											"Данные ребёнка (детей) /лица из числа детей-сирот и детей, оставшихся без попечения родителей: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет"))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("отказ в предоставлении услуг отдыха и оздоровления."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание отказа: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(_FEDERAL_LAW_REF))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Причина отказа: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(_PARTICIPATE_NOTIFICATION))));


					SignBlockNotification2020(doc, account, "Исполнитель:");
					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в предоставлении услуг отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuseCompensation(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text(
								"Уведомление об отказе в выплате компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата обращения: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"отказ в предоставлении выплаты компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(_FEDERAL_LAW))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Причина отказа в выплате: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(request.DeclineReason?.Name))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					SignBlockNotification(doc, account,
						$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuseCompensationYouthRest(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text(
								"Уведомление об отказе в выплате компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата обращения: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"отказ в предоставлении выплаты компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(_FEDERAL_LAW))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Причина отказа в выплате: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(request.DeclineReason?.Name))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					SignBlockNotification(doc, account,
						$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuse1080(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление об отказе в предоставлении услуг отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));


					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text(
												"Данные ребенка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Льготная категория: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{child.BenefitType?.Name}"))));
						}
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text(
											"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет"))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("отказ в предоставлении услуг отдыха и оздоровления."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание отказа: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(_FEDERAL_LAW))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Причина отказа: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(request.DeclineReason?.Name))));


					SignBlockNotification2020(doc, account, "Исполнитель:");
					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в предоставлении услуг отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationRefuseContent(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				string notificationName;
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					notificationName = "Уведомление о рассмотрении заявления";

					if (request.StatusId == (long)StatusEnum.CancelByApplicant)
					{
						notificationName =
							"Уведомление о прекращении рассмотрения поданного заявления о предоставлении услуг отдыха и оздоровления";
					}
					else if (request.StatusId == (long)StatusEnum.Reject)
					{
						notificationName = request.TypeOfRestId == (long)TypeOfRestEnum.Compensation ||
										   request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest
							? "Уведомление об отказе в выплате компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления"
							: "Уведомление об отказе в предоставлении услуг отдыха и оздоровления";
					}
					else if (request.StatusId == (long)StatusEnum.CertificateIssued)
					{
						notificationName = request.TypeOfRestId == (long)TypeOfRestEnum.Compensation ||
										   request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest
							? "Уведомление о предоставлении выплаты компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления"
							: "Уведомление о неучастии заявителя в выборе конкретной организации отдыха и оздоровления";
					}

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text(notificationName))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);

					var applicant = request.NullSafe(r => r.Applicant) ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					if ((request.TypeOfRestId == (long)TypeOfRestEnum.Compensation ||
						 request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest) &&
						request.StatusId == (long)StatusEnum.Reject)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Дата обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.DateRequest?.Date.FormatEx()))));
					}
					else
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Дата и время регистрации заявления: ")
									{
										Space = SpaceProcessingModeValues.Preserve
									}),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.DateRequest.FormatEx()))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatEx(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					if (!request.IsFirstCompany)
					{
						if (request.StatusId == (long)StatusEnum.CancelByApplicant)
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Причина обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											"прекращение рассмотрения поданного заявления о предоставлении услуг отдыха и оздоровления по инициативе заявителя."))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Результат рассмотрения заявления: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text("отозвано по инициативе заявителя в установленный Порядком срок."))));
						}
						else
						{
							if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation &&
								request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest)
							{
								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Наименование запрашиваемой услуги:  ")
											{
												Space = SpaceProcessingModeValues.Preserve
											}),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text("___________________________________________"))));

								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Цель обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text(request.TypeOfRest?.Name))));

								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Организация отдыха и оздоровления: ")
											{
												Space = SpaceProcessingModeValues.Preserve
											}),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text(request.Hotels?.Name ??
													 request.Tour?.Hotels?.Name ?? string.Empty))));
							}

							if (request.Tour != null)
							{
								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Время отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text(request.Tour.DateIncome.FormatEx() + " - " +
													 request.Tour.DateOutcome.FormatEx()))));
							}

							if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation
								&& request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest ||
								request.StatusId != (long)StatusEnum.Reject)
							{
								AppendChildrenToDocument(doc, request);
							}

							if ((request.TypeOfRestId == (long)TypeOfRestEnum.Compensation
								 || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest) &&
								request.StatusId == (long)StatusEnum.Reject)
							{
								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Both },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Причина обращения: ")
											{ Space = SpaceProcessingModeValues.Preserve }),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text(
												"выплата компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку."))));
							}

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Результат рассмотрения заявления: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true), new Text(
										(request.TypeOfRestId == (long)TypeOfRestEnum.Compensation
										 || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest) &&
										request.StatusId == (long)StatusEnum.Reject
											? "отказ в предоставлении выплаты компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления."
											: request.Status.Name))));
						}


						if ((request.TypeOfRestId == (long)TypeOfRestEnum.Compensation
							 || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest) &&
							(request.StatusId == (long)StatusEnum.CertificateIssued ||
							 request.StatusId == (long)StatusEnum.Reject))
						{
							if (request.DeclineReason != null)
							{
								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Both },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Причина отказа в выплате: ")
											{ Space = SpaceProcessingModeValues.Preserve }),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true), new Text(
											request.DeclineReason.Name))));
							}

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunProperties.CloneNode(true),
										new Text(_FEDERAL_LAW))));
						}
						else if (request.DeclineReason != null)
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										//new Text((request.TypeOfRestId == (long) TypeOfRestEnum.Compensation
										//	         ? "приказ Департамента культуры города Москвы от 12 января 2016 г. № 8 \"О выплате частичной компенсации стоимости самостоятельно приобретенной путевки на отдых и оздоровление детей и сопровождающих их лиц в 2016 году\", пункт "
										//	         : "") + request.DeclineReason.Name)
										new Text(_FEDERAL_LAW))));
						}
					}
					else if (request.StatusId == (long)StatusEnum.CancelByApplicant)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Причина обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										"прекращение рассмотрения поданного заявления о предоставлении услуг отдыха и оздоровления по инициативе заявителя."))));
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения заявления: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text("отозвано по инициативе заявителя в установленный Порядком срок"))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(_FEDERAL_LAW))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					}
					else if (request.StatusId == (long)StatusEnum.Reject)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.TypeOfRest?.Name))));

						AppendChildrenToDocument(doc, request);

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения заявления: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text("отказ в предоставлении услуг отдыха и оздоровления"))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Причина отказа: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.DeclineReason?.Name ?? "-"))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(_FEDERAL_LAW))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					}
					else if (request.StatusId == (long)StatusEnum.CertificateIssued)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.TypeOfRest?.Name))));

						AppendChildrenToDocument(doc, request);

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения заявления: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true), new Text("Услуга оказана. Решение отрицательное"))));

						//doc.AppendChild(
						//    new Paragraph(
						//        new ParagraphProperties(new Justification {Val = JustificationValues.Left},
						//            new SpacingBetweenLines {After = _SIZE_20}),
						//        new Run(titleRequestRunPropertiesBold.CloneNode(true),
						//            new Text("Номер путевки: ") {Space = SpaceProcessingModeValues.Preserve}),
						//        new Run(titleRequestRunPropertiesItalic.CloneNode(true),
						//            new Text(request.CertificateNumber))));

						//doc.AppendChild(
						//    new Paragraph(
						//        new ParagraphProperties(new Justification {Val = JustificationValues.Left},
						//            new SpacingBetweenLines {After = _SIZE_20}),
						//        new Run(titleRequestRunPropertiesBold.CloneNode(true),
						//            new Text("Организация отдыха и оздоровления: ")
						//                {Space = SpaceProcessingModeValues.Preserve}),
						//        new Run(titleRequestRunPropertiesItalic.CloneNode(true),
						//            new Text(request.Hotels?.Name ?? request.Tour?.Hotels?.Name ?? string.Empty))));

						//doc.AppendChild(
						//    new Paragraph(
						//        new ParagraphProperties(new Justification {Val = JustificationValues.Left},
						//            new SpacingBetweenLines {After = _SIZE_20}),
						//        new Run(titleRequestRunPropertiesBold.CloneNode(true),
						//            new Text("Время отдыха: ") {Space = SpaceProcessingModeValues.Preserve}),
						//        new Run(titleRequestRunPropertiesItalic.CloneNode(true),
						//            new Text(request.TimeOfRest?.Name ?? string.Empty))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(_DECLINE_REASON_PARTICIPATE))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text("    " + _PARTICIPATE) { Space = SpaceProcessingModeValues.Preserve })));

					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					if (request.StatusId == (long)StatusEnum.CertificateIssued)
					{
						SignBlockNotification2020(doc, account, "Исполнитель");
					}
					else
					{
						SignBlockNotification(doc, account,
											$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}",
											request.StatusId != (long)StatusEnum.CertificateIssued);
					}


					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = $"{notificationName}" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}
		#endregion

		protected override IDocument NotificationWaitApplicant(Request request, Account account, IEnumerable<BenefitType> benefits)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(
								new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление о приостановлении рассмотрения заявления"),
							new Break(),
							new Text("о предоставлении услуг отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = new RunProperties().SetFont().SetFontSize();
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));

					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text(
											"Данные ребенка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text($"{child.BenefitType?.Name}"))));
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text(
											"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет"))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Этап рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("рассмотрение заявления о предоставлении услуг отдыха")
								{ Space = SpaceProcessingModeValues.Preserve },
								new Break(),
								new Text("и оздоровления приостановлено.") { Space = SpaceProcessingModeValues.Preserve }
							)));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("пункт 6.4 Порядка организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденного постановлением Правительства Москвы") { Space = SpaceProcessingModeValues.Preserve },
								new Break(),
								new Text("от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\": ") { Space = SpaceProcessingModeValues.Preserve }),
							// new Break(),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("\"Необходимость личной явки заявителя в ГАУК \"МОСГОРТУР\".") { Space = SpaceProcessingModeValues.Preserve }
								)));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										"Для подтверждения сведений, указанных в заявлении о предоставлении услуг отдыха")
								{ Space = SpaceProcessingModeValues.Preserve },
								new Break(),
								new Text(" и оздоровления (далее – заявление), ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("в течение 10 рабочих дней со дня направления данного уведомления")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										" Вам необходимо явиться в офис ГАУК \"МОСГОРТУР\" по адресу: г. Москва, Малый Харитоньевский переулок, дом 6 строение 3.")
								{ Space = SpaceProcessingModeValues.Preserve })));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Прием в офисе ГАУК \"МОСГОРТУР\" осуществляется исключительно")
								{ Space = SpaceProcessingModeValues.Preserve },
								new Break(),
								new Text("по предварительной записи. ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("Запись производится через официальный портал Мэра")
								{ Space = SpaceProcessingModeValues.Preserve },
								new Break(),
								new Text(
										"и Правительства Москвы mos.ru или при личном визите заявителя в офис ГАУК \"МОСГОРТУР\". Запись производится на свободную дату и время.")
								{ Space = SpaceProcessingModeValues.Preserve }
							)));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("При себе необходимо иметь: ") { Space = SpaceProcessingModeValues.Preserve })));


					//List<string> docs = new List<string>
					//{
					//    "документ, удостоверяющий личность заявителя;",
					//    "документы, подтверждающие, что заявитель является родителем ребёнка (свидетельство о рождении ребенка, в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения);",
					//    "документ, подтверждающий место жительства ребёнка, лица из числа детей-сирот и детей, оставшихся без попечения родителей, в городе Москве;",
					//    "документ, подтверждающий полномочия заявителя, сопровождающего лица (в случае организации совместного выездного отдыха) из числа законных представителей – родителей, опекунов, попечителей, приемных родителей, патронатных воспитателей ребёнка (договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие статус ребенка);",
					//   "документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
					//    "документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребёнка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);",
					//    "документ, подтверждающий отнесение ребёнка, к одной из категорий детей, находящихся в трудной жизненной ситуации и указанных в пунктах 3.1.3, 3.1.5 - 3.1.13 Порядка, лица из числа детей-сирот к категории лиц из числа детей-сирот и детей, оставшихся без попечения родителей (заключение медико-социальной экспертизы, заключение Центральной психолого-медико-педагогической комиссии города Москвы, справки уполномоченного учреждения социальной защиты населения города Москвы и/или федеральных органов);"
					//};


					//доработка для наркоманов
					List<string> innerListOrphans = new List<string>();
					List<string> innerListDisabled = new List<string>();
					List<string> innerListLowIncome = new List<string>();
					List<string> innerListSacrifice = new List<string>();
					List<string> innerListRefugee = new List<string>();
					List<string> innerListExtreme = new List<string>();
					List<string> innerListViolence = new List<string>();
					List<string> innerListInvalid = new List<string>();
					List<string> innerListTerror = new List<string>();
					List<string> innerListMilitary = new List<string>();
					List<string> innerListInvalidParents = new List<string>();
					List<string> innerListDeviant = new List<string>();
					List<string> innerListOrphansYouth = new List<string>();


					if (benefits.Count() > 0 || (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest))
					{


						foreach (BenefitType benefit in benefits)
						{
							if (benefit.ExnternalUid.Contains("52"))//дети-сироты и оставшиеся без попечения родителей...
							{
								innerListOrphans = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является законным представителем ребенка: договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие полномочия законного представителя ребенка;",
										"документ, удостоверяющий личность ребенка;",
										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий полномочия заявителя, сопровождающего лица (в случае организации совместного выездного отдыха) из числа законных представителей – опекунов, попечителей, приемных родителей, патронатных воспитателей ребенка (договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие статус ребенка);",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность).",

									};


							}
							if (benefit.ExnternalUid.Contains("24"))//дети-инвалиды, дети с ограниченными возможностями...
							{

								innerListDisabled = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем (законным представителем) ребенка: свидетельство о рождении ребенка*, договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие полномочия законного представителя ребенка;",
										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, удостоверяющий личность ребенка;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении (заключение медико-социальной экспертизы или заключение Центральной психолого-медико-педагогической комиссии города Москвы**);",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени;",
										"** является документом, подтверждающим категорию, только в случае указания на необходимость предоставления ребенку специальных условий обучения по адаптированной основной образовательной программе."
									};



							}
							if (benefit.ExnternalUid.Contains("48"))//дети из малообеспеченных семей
							{

								innerListLowIncome = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, удостоверяющий личность ребенка;",
										"документ, подтверждающий место жительства ребенка в городе Москве;",

										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}

							if (benefit.ExnternalUid.Contains("57,71,72"))
							{
								innerListSacrifice = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 7)
							{
								innerListRefugee = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 8)
							{
								innerListExtreme = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 9)
							{
								innerListViolence = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 10)
							{
								innerListInvalid = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 11)
							{
								innerListTerror = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.ExnternalUid.Contains("58,71,72"))
							{
								innerListMilitary = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.ExnternalUid.Contains("56"))
							{
								innerListInvalidParents = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}
							if (benefit.BenefitTypeERLId == 14)
							{
								innerListDeviant = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документы, подтверждающие, что заявитель является родителем ребенка – свидетельство о рождении ребенка*;",

										"документ, подтверждающий место жительства ребенка в городе Москве;",
										"документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;",
										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени."

									};
							}

						}

						if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest)//лица из числа детей-сирот и детей, оставших без попечения...
						{


							innerListOrphansYouth = new List<string>
									{
										"документ, удостоверяющий личность заявителя;",
										"документ, подтверждающий место жительства лица из числа детей-сирот и детей, оставшихся без попечения родителей, в городе Москве;",

										"документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);",
										"документ, подтверждающий отнесение лица из числа детей-сирот и детей, оставшихся без попечения родителей, к категории лиц из числа детей-сирот и детей, оставшихся без попечения родителей."

									};


						}



					}

					//if (benefits.Count() > 0 ||(request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest))
					//{
					//    docs.Clear();
					//    docs.Add("документ, удостоверяющий личность заявителя;");

					//    docs.Add("документ, подтверждающий место жительства ребенка в городе Москве;");
					//    docs.Add("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//    docs.Add("документы, подтверждающие, что заявитель является родителем (законным представителем) ребенка: свидетельство о рождении ребенка*, договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие полномочия законного представителя ребенка;");
					//    docs.Add("документ, подтверждающий полномочия доверенного лица на совершение действий в период проведения заявочной кампании (в случае подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала доверенным лицом на совершение действий в период проведения заявочной кампании) (нотариально заверенное согласие или доверенность);");

					//    docs.Add("* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени");
					//    foreach (BenefitType benefit in benefits)
					//    {
					//        if (benefit.ExnternalUid.Contains("52"))//дети-сироты и оставшиеся без попечения родителей...
					//        {
					//            List<string> innerList = new List<string>
					//            {
					//                "документ, удостоверяющий личность ребенка;",
					//                "документ, подтверждающий полномочия заявителя, сопровождающего лица(в случае организации совместного выездного отдыха) из числа законных представителей – опекунов, попечителей, приемных родителей, патронатных воспитателей ребенка(договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие статус ребенка);",
					//                "документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);",

					//            };

					//            docs.InsertRange(4,innerList);
					//            docs.Remove("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//            docs.Remove("* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени");
					//        }
					//        if (benefit.ExnternalUid.Contains("24"))//дети-инвалиды, дети с ограниченными возможностями...
					//        {
					//            docs.Insert(1,"документ, удостоверяющий личность ребенка;");
					//            docs.Insert(3, "документ, подтверждающий отнесение ребенка к категории, указанной в заявлении (заключение медико-социальной экспертизы или заключение Центральной психолого-медико-педагогической комиссии города Москвы**);");
					//            docs.Insert(5,"документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);");
					//            docs.Remove("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//            docs.Add("** является документом, подтвержающим категорию, только в случае указания на необходимость предоставления ребенку специальных условий обучения по адаптированной основной образовательной программе");
					//        }
					//        if (benefit.ExnternalUid.Contains("48"))//дети из малообеспеченных семей
					//        {

					//            docs.Insert(1,"документ, удостоверяющий личность ребенка;");
					//            docs.Insert(5,"документ, подтверждающий полномочия доверенного лица для сопровождения во время отдыха и оздоровления (в случае организации совместного выездного отдыха и сопровождения ребенка доверенным лицом для сопровождения во время отдыха и оздоровления и подачи заявления о предоставлении услуг отдыха и оздоровления с использованием Портала) (нотариально заверенное согласие или доверенность);");
					//            docs.Remove("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//        }


					//    }

					//    if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps || request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest)//лица из числа детей-сирот и детей, оставших без попечения...
					//    {

					//        docs.Insert(1, "документ, подтверждающий место жительства лица из числа детей-сирот и детей, оставшихся без попечения родителей, в городе Москве;");
					//        docs.Insert(2, "документ, подтверждающий отнесение лица из числа детей-сирот и детей, оставшихся без попечения родителей, к категории лиц из числа детей-сирот и детей, оставшихся без попечения родителей.");
					//        docs.Remove("документ, подтверждающий место жительства ребенка в городе Москве;");
					//        docs.Remove("документы, подтверждающие, что заявитель является родителем (законным представителем) ребенка: свидетельство о рождении ребенка*, договор о приемной семье, распоряжение об опеке, иные документы, устанавливающие полномочия законного представителя ребенка;");
					//        docs.Remove("документ, подтверждающий отнесение ребенка к категории, указанной в заявлении;");
					//        docs.Remove("* в случае если с момента рождения ребенка у родителя произошла смена фамилии, имени или отчества – необходимо также предоставить документы, подтверждающие данные изменения: свидетельство о браке, свидетельство о расторжении брака, свидетельство о перемене имени");
					//    }
					//}


					//foreach (var docText in docs)
					//{
					//    doc.AppendChild(
					//        new Paragraph(
					//            new ParagraphProperties(new Justification {Val = JustificationValues.Both},
					//                new SpacingBetweenLines {After = _SIZE_20},
					//                new Indentation {FirstLine = _FIRST_LINE_INDENTATION_600.ToString()}),
					//            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
					//                new Text(docText)
					//                    {Space = SpaceProcessingModeValues.Preserve})));
					//}
					if (innerListOrphans.Count > 0)
					{
						doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Дети-сироты и дети, оставшиеся без попечения родителей, находящиеся под опекой, попечительством, в том числе в приёмной или патронатной семье")
										{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListOrphans)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListDisabled.Count > 0)
					{

						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети-инвалиды и дети с ограниченными возможностями здоровья")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListDisabled)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListLowIncome.Count > 0)
					{

						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети из малообеспеченных семей")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListLowIncome)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListSacrifice.Count > 0)
					{

						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети - жертвы вооруженных и межнациональных конфликтов, экологических и техногенных катастроф, стихийных бедствий")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListSacrifice)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListRefugee.Count > 0)
					{

						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети из семей беженцев и вынужденных переселенцев")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListRefugee)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListExtreme.Count > 0)
					{
						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети, оказавшиеся в экстремальных условиях")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListExtreme)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListViolence.Count > 0)
					{
						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети - жертвы насилия")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListViolence)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListInvalid.Count > 0)
					{
						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети, жизнедеятельность которых объективно нарушена в результате сложившихся обстоятельств и которые не могут преодолеть данные обстоятельства самостоятельно или с помощью семьи")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListInvalid)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListTerror.Count > 0)
					{
						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети, пострадавшие в результате террористических актов")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListTerror)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListMilitary.Count > 0)
					{
						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети из семей военнослужащих и приравненных к ним лиц, погибших или получивших увечья (ранения, травмы, контузии) при исполнении ими обязанностей военной службы или служебных обязанностей")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListMilitary)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListInvalidParents.Count > 0)
					{
						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети из семей, в которых оба или один родитель являются инвалидами")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListInvalidParents)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListDeviant.Count > 0)
					{
						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Дети с отклонениями в поведении")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListDeviant)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					if (innerListOrphansYouth.Count > 0)
					{
						doc.AppendChild(
	new Paragraph(
		new ParagraphProperties(new Justification { Val = JustificationValues.Both },
			new SpacingBetweenLines { After = _SIZE_20 },
			new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
		new Run(titleRequestRunPropertiesBold.CloneNode(true),
			new Text("Лица из числа детей-сирот и детей, оставшихся без попечения родителей")
			{ Space = SpaceProcessingModeValues.Preserve })));
						foreach (var docText in innerListOrphansYouth)
						{


							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Both },
										new SpacingBetweenLines { After = _SIZE_20 },
										new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(docText)
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										"Неявка с соответствующими документами в течение срока приостановления рассмотрения заявления является основанием для отказа в предоставлении услуг отдыха и оздоровления.")
								{ Space = SpaceProcessingModeValues.Preserve })));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));


					SignBlockNotification2020(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о приостановлении рассмотрения заявления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		#region NotificationAboutDecision

		protected override IDocument NotificationAboutCompensationIssued(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var isCompensationYouth = request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest;

					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					if (isCompensationYouth)
					{
						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text("Уведомление о предоставлении выплаты компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления"))));
					}
					else
					{
						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text("Уведомление о предоставлении выплаты компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления"))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text(
												$"{(isCompensationYouth ? "Данные лица из числа детей-сирот и детей, оставшихся без попечения родителей" : "Данные о ребёнке")}: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text(
												$"{(isCompensationYouth ? "Льготная категория" : "Льготная категория ребёнка")}: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{child.BenefitType?.Name}"))));
						}
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("Услуга оказана."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\"."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					SignWorkerCompensationBlock(doc, account);

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о предоставлении выплаты компенсации за самостоятельно приобретенную путевку" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationAboutCertificate(Request request, Account account)
		{

			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление о предоставлении сертификата на отдых и оздоровление"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name ?? "сертификат на отдых и оздоровление"))));


					if (request.Child != null)
					{
						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Данные лиц, указанных в сертификате: ") { Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Льготная категория ребенка: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{child.BenefitType?.Name}"))));
						}
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("Услуга оказана."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер сертификата: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.CertificateNumber))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									"Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\"."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));


					SignBlockNotification2020(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о предоставлении сертификата на отдых и оздоровление" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument NotificationAboutTour(Request request, Account account, bool forPrint = false)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление о предоставлении бесплатной путевки для отдыха"),
							new Break(),
							new Text("и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));

					if (request.Child != null && request.Child.Count > 0)
					{
						if (request.Applicant.IsAccomp || request.Attendant.Any())
						{
							doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text(
													"Данные лиц, указанных в путевке: ")
											{ Space = SpaceProcessingModeValues.Preserve })));
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text(
												"Сопровождающие: ")
										{ Space = SpaceProcessingModeValues.Preserve })));

							if (request.Applicant.IsAccomp)
							{
								doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{request.Applicant.LastName} {request.Applicant.FirstName} {request.Applicant.MiddleName}, {request.Applicant.DateOfBirth.FormatExGR(string.Empty)}"))));
							}

							if (request.Attendant.Any())
							{
								foreach (var attendant in request.Attendant.Where(c => !c.IsDeleted))
								{
									doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text(
											$"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}, {attendant.DateOfBirth.FormatExGR(string.Empty)}"))));
								}
							}
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text(
												"Дети: ")
										{ Space = SpaceProcessingModeValues.Preserve })));
						}

						foreach (var child in request.Child.Where(c => !c.IsDeleted))
						{
							if (request.Child.Count == 1 && request.TypeOfRest.ParentId != 2 && request.TypeOfRest.ParentId != 9)
							{
								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text(
													"Данные лиц, указанных в путевке: ")
											{ Space = SpaceProcessingModeValues.Preserve }),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text(
												$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Льготная категория: ")
											{ Space = SpaceProcessingModeValues.Preserve }),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text($"{child.BenefitType?.Name}"))));
							}
							else
							{
								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										//new Run(titleRequestRunPropertiesBold.CloneNode(true),
										//    new Text(
										//            "Данные лиц, указанных в путевке ")
										//        {Space = SpaceProcessingModeValues.Preserve}),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text(
												$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

								doc.AppendChild(
									new Paragraph(
										new ParagraphProperties(new Justification { Val = JustificationValues.Left },
											new SpacingBetweenLines { After = _SIZE_20 }),
										new Run(titleRequestRunPropertiesBold.CloneNode(true),
											new Text("Льготная категория: ")
											{ Space = SpaceProcessingModeValues.Preserve }),
										new Run(titleRequestRunPropertiesItalic.CloneNode(true),
											new Text($"{child.BenefitType?.Name}"))));
							}
						}
					}

					if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text(
											"Данные лиц, указанных в путевке: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										"Лица из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет"))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("Услуга оказана."))));

					/*doc.AppendChild(
                        new Paragraph(
                            new ParagraphProperties(new Justification {Val = JustificationValues.Left},
                                new SpacingBetweenLines {After = _SIZE_20}),
                            new Run(titleRequestRunPropertiesBold.CloneNode(true),
                                new Text("Дата и время выбора варианта организации отдыха и оздоровления: ")
                                    {Space = SpaceProcessingModeValues.Preserve}),
                            new Run(titleRequestRunPropertiesItalic.CloneNode(true),
                                new Text($"{request.DateChangeStatus:dd.MM.yyyy HH:mm}"))));*/

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер путевки: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.CertificateNumber))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Организация отдыха и оздоровления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.Tour?.Hotels?.Name))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Время отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text($"{request.TimeOfRest?.Name} ({request.Tour?.DateIncome.FormatEx()} - {request.Tour?.DateOutcome.FormatEx()})"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text($"{_FEDERAL_SHORT2021_LAW}."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));


					SignBlockNotification2020(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о предоставлении бесплатной путевки для отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}
		#endregion

		public override IDocument NotificationOrgChoose(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(
								new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление о необходимости выбора организации отдыха"),
							new Break(),
							new Text("и оздоровления")
						)));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());
					var titleRequestRunPropertiesUnderline = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesUnderline.AppendChild(new Underline { Val = UnderlineValues.Single });

					var applicant = request.Applicant ?? new Applicant
					{ DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center },
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text($"Уважаемый(ая) {applicant.LastName} {applicant.FirstName} {applicant.MiddleName},")
								{ Space = SpaceProcessingModeValues.Preserve }))));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text($"Ваше заявление от {request.DateRequest?.Date.FormatEx()} г. № {request.RequestNumber} о предоставлении услуг отдыха и оздоровления (далее – заявление) рассмотрено.")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("В период второго этапа заявочной кампании (с 7 по 21 февраля 2023 г.) Вам необходимо дополнить Ваше заявление сведениями о конкретной организации отдыха и оздоровления. Выбор конкретной организации отдыха и оздоровления осуществляется из числа предлагаемых ГАУК\u00A0\"МОСГОРТУР\" в соответствии с указанными Вами на первом этапе заявочной кампании сведениями о приоритетном времени, направлении отдыха и оздоровления и количестве детей.")
								{ Space = SpaceProcessingModeValues.Preserve })
						));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("В случае подачи заявления через подсистему \"личный кабинет\" Портала Мэра и Правительства Москвы ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesUnderline.CloneNode(true),
								new Text("mos.ru")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(" (далее – подсистема \"личный кабинет\" Портала), дополнение формы заявления сведениями о конкретной организации отдыха и оздоровления осуществляется в подсистеме \"личный кабинет\" Портала.")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("В случае подачи заявления при личном обращении в офисе ГАУК \"МОСГОРТУР\", дополнение формы заявления сведениями о конкретной организации отдыха и оздоровления возможно только при личном обращении заявителя в офис ГАУК \"МОСГОРТУР\" по адресу: г.\u00A0Москва, Малый Харитоньевский переулок д. 6, стр. 3.")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Прием в офисе ГАУК \"МОСГОРТУР\" осуществляется исключительно по предварительной записи.")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(" Запись производится через Портал Мэра и Правительства Москвы mos.ru (далее – Портала) или при личном визите заявителя в офис ГАУК \"МОСГОРТУР\". Запись производится на свободную дату и время.")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("В случае если на втором этапе заявочной кампании Вас не устроит ни один из вариантов организаций отдыха и оздоровления, предлагаемых ГАУК \"МОСГОРТУР\", на втором этапе заявочной кампании в период с 7 по 21 февраля 2023 г. Вам необходимо отказаться от всех предложенных организаций отдыха и оздоровления.")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("При подаче заявления через подсистему \"личный кабинет\" Портала для отказа от всех предложенных организаций отдыха и оздоровления необходимо нажать соответствующую \"галочку\" в интерактивном поле Портала: \"Я отказываюсь от предложенных вариантов организаций отдыха и оздоровления\".")
								{ Space = SpaceProcessingModeValues.Preserve })
						));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("При подаче заявления в офисе ГАУК \"МОСГОРТУР\" для отказа от всех предложенных организаций отдыха и оздоровления необходимо лично обратиться в офис ГАУК \"МОСГОРТУР\".")
								{ Space = SpaceProcessingModeValues.Preserve })
						));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));


					SignWorkerBlock(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о необходимости выбора организации отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		protected override IDocument InnerNotificationAttendantChange(Request request, Account account, long oldAttendantId = 0, long attendantId = 0, string changeDate = null)
		{
			if (oldAttendantId == 0)
				oldAttendantId = request.Attendant.Where(att => att.IsDeleted).Any() ? request.Attendant.Where(att => att.IsDeleted).OrderByDescending(att => att.Id).ToList().FirstOrDefault().Id : request.Applicant.Id;
			if (attendantId == 0)
				attendantId = request.Attendant.Where(att => !att.IsDeleted).OrderByDescending(att => att.Id).ToList().FirstOrDefault().Id;

			var isCert = request.TypeOfRestId == (long)TypeOfRestEnum.Money || request.TypeOfRest.ParentId == (long)TypeOfRestEnum.Money;
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));



					var elems = new List<OpenXmlElement>();
					elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold());
					if (isCert)
						elems.Add(new Text("Уведомление о замене сопровождающего лица в предоставленном сертификате на отдых и оздоровление"));
					else
						elems.Add(new Text("Уведомление о замене сопровождающего лица в предоставленной бесплатной путевке для отдыха и оздоровления"));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(elems)
					));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var titleRequestRunPropertiesUnderline = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesUnderline.AppendChild(new Indentation());

					var documentRunPropertiesItalic = new RunProperties().SetFont().SetFontSize(_SIZE_16);
					documentRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };
					Applicant oldAttenadant = null;
					Applicant attendant = request.Attendant.FirstOrDefault(att => att.Id == attendantId);
					if (request.Applicant.Id == oldAttendantId)
						oldAttenadant = applicant;
					else oldAttenadant = request.Attendant.FirstOrDefault(att => att.Id == oldAttendantId);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					if (isCert)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер сертификата: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.CertificateNumber.FormatEx()))));
					}
					else
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер путевки: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.CertificateNumber.FormatEx()))));
					}

					doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.RequestNumber.FormatEx()))));

					if (!request.Tour.IsNullOrEmpty())
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Организация отдыха и оздоровления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.Tour.Hotels.NameOrganization ?? request.Tour.Hotels.Name))));

					if (!request.Tour.IsNullOrEmpty())
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Время отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.Tour.DateIncome.FormatEx("dd.MM.yyyy") + " - " + request.Tour.DateOutcome.FormatEx("dd.MM.yyyy")))));

					//doc.AppendChild(
					//    new Paragraph(
					//        new ParagraphProperties(new Justification { Val = JustificationValues.Left },
					//            new SpacingBetweenLines { After = _SIZE_20 }),
					//        new Run(titleRequestRunPropertiesBold.CloneNode(true),
					//            new Text("Дата замены сопровождающего лица: ")
					//            { Space = SpaceProcessingModeValues.Preserve }),
					//        new Run(titleRequestRunPropertiesItalic.CloneNode(true),

					//            new Text(System.DateTime.Now.FormatEx("dd.MM.yyyy")))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные сопровождающего лица, указанного при подаче заявления о предоставлении услуг отдыха и оздоровления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{oldAttenadant.LastName} {oldAttenadant.FirstName} {oldAttenadant.MiddleName}, {oldAttenadant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные предлагаемого к замене сопровождающего лица, указанного в заявлении о замене сопровождающего лица: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{attendant.LastName} {attendant.FirstName} {attendant.MiddleName}, {attendant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание приема заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\"."))));

					if (isCert)
					{
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("Ваше заявление о замене сопровождающего лица в предоставленном сертификате на отдых и оздоровление рассмотрено. Сопровождающее лицо изменено. Сертификат на отдых и оздоровление с измененными сведениями о сопровождающем лице прилагается."))
							));
					}
					else
					{
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("Ваше заявление о замене сопровождающего лица в предоставленной путевке для отдыха и оздоровления рассмотрено положительно. Путевка для отдыха и оздоровления с измененными сведениями о сопровождающем лице прилагается."))
							));
					}

					if (isCert)
					{
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата написания заявления о замене сопровождающего лица в предоставленном сертификате на отдых и оздоровление: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
							new Text(oldAttenadant.AttendantChangeRequestDate.HasValue ? oldAttenadant.AttendantChangeRequestDate.FormatEx("dd.MM.yyyy") : changeDate.IsNullOrEmpty() ? System.DateTime.Now.FormatEx("dd.MM.yyyy") : changeDate))
							//new Text(changeDate.IsNullOrEmpty() ? System.DateTime.Now.FormatEx("dd.MM.yyyy") : changeDate))
							));
						//doc.AppendChild(
						//new Paragraph(
						//    new ParagraphProperties(new Justification { Val = JustificationValues.Both },
						//        new SpacingBetweenLines { After = _SIZE_20 }),
						//    new Run(titleRequestRunPropertiesBold.CloneNode(true),
						//        new Text("Дата замены сопровождающего лица в предоставленном сертификате на отдых и оздоровление: ")
						//        { Space = SpaceProcessingModeValues.Preserve }),
						//    new Run(titleRequestRunPropertiesItalic.CloneNode(true),
						//    new Text(oldAttenadant.AttendantChangeDate.HasValue ? oldAttenadant.AttendantChangeDate.FormatEx("dd.MM.yyyy") : changeDate.IsNullOrEmpty() ? System.DateTime.Now.FormatEx("dd.MM.yyyy") : changeDate))
						//    //new Text(changeDate.IsNullOrEmpty() ? System.DateTime.Now.FormatEx("dd.MM.yyyy") : changeDate))
						//    ));
					}
					else
					{
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата написания заявления о замене сопровождающего лица в предоставленной бесплатной путевке для отдыха и оздоровления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(oldAttenadant.AttendantChangeRequestDate.HasValue ? oldAttenadant.AttendantChangeRequestDate.FormatEx("dd.MM.yyyy") : changeDate.IsNullOrEmpty() ? System.DateTime.Now.FormatEx("dd.MM.yyyy") : changeDate))
							));
						//doc.AppendChild(
						//new Paragraph(
						//    new ParagraphProperties(new Justification { Val = JustificationValues.Both },
						//        new SpacingBetweenLines { After = _SIZE_20 }),
						//    new Run(titleRequestRunPropertiesBold.CloneNode(true),
						//        new Text("Дата замены сопровождающего лица в предоставленной бесплатной путевке для отдыха и оздоровления: ")
						//        { Space = SpaceProcessingModeValues.Preserve }),
						//    new Run(titleRequestRunPropertiesItalic.CloneNode(true),
						//        new Text(oldAttenadant.AttendantChangeDate.HasValue ? oldAttenadant.AttendantChangeDate.FormatEx("dd.MM.yyyy") : changeDate.IsNullOrEmpty() ? System.DateTime.Now.FormatEx("dd.MM.yyyy") : changeDate))
						//    ));
					}


					doc.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Both },
				new SpacingBetweenLines { After = _SIZE_20 })));


					//if (isCert)
					//{
					//    doc.AppendChild(
					//    new Paragraph(
					//        new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//            new SpacingBetweenLines { After = _SIZE_20 },
					//            new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
					//        new Run(titleRequestRunProperties.CloneNode(true),
					//            new Text("Сертификат на отдых и оздоровление с измененными сведениями о сопровождающем лице прилагается.")
					//            { Space = SpaceProcessingModeValues.Preserve })));
					//}
					//else
					//{
					//    doc.AppendChild(
					//    new Paragraph(
					//        new ParagraphProperties(new Justification { Val = JustificationValues.Both },
					//            new SpacingBetweenLines { After = _SIZE_20 },
					//            new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
					//        new Run(titleRequestRunProperties.CloneNode(true),
					//            new Text("Путевка для отдыха и оздоровления с измененными сведениями о сопровождающем лице прилагается.")
					//            { Space = SpaceProcessingModeValues.Preserve })));
					//}

					if (isCert)
					{
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("Приложение: сертификат на отдых и оздоровление.")
								{ Space = SpaceProcessingModeValues.Preserve })));
					}
					else
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("Приложение: путевка для отдыха и оздоровления.")
									{ Space = SpaceProcessingModeValues.Preserve })));
					}

					doc.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Both },
				new SpacingBetweenLines { After = _SIZE_20 })));



					SignBlockNotification2022(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о замене сопровождающего лица" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		public override IDocument NotificationDeclineRegistryReg(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var isCompensationYouth = request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest;

					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление об отказе в регистрации заявления о предоставлении услуг отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Bold(), new Text(_SPACE),
								new Text("Номер отказа: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Italic(), new Text(_SPACE),
								new Text(request.RequestNumber?.FormatEx() ?? "-"))
							));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Bold(), new Text(_SPACE),
								new Text("Дата обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Italic(), new Text(_SPACE),
								new Text(request.DateRequest?.FormatEx() ?? "-"))
							));

					string applicantBirthDate = request.Applicant?.DateOfBirth?.ToString("dd.MM.yyyy");
					string applicantData = $"{request.Applicant?.LastName ?? string.Empty} {request.Applicant?.FirstName ?? string.Empty} {request.Applicant?.MiddleName ?? string.Empty}";
					applicantData += !string.IsNullOrWhiteSpace(applicantBirthDate) ? $", {applicantBirthDate} г.р." : string.Empty;

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Bold(), new Text(_SPACE),
								new Text("Данные обратившегося лица: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Italic(), new Text(_SPACE),
								new Text(applicantData))
							));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Bold(), new Text(_SPACE),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Italic(), new Text(_SPACE),
								new Text(((request.Applicant?.Phone ?? "-")
									+ ", " + request.Applicant?.Email ?? "-")))
							));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Bold(), new Text(_SPACE),
								new Text("Цель обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Italic(), new Text(_SPACE),
								new Text((request.TypeOfRest?.Parent?.Name)))
							));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Bold(), new Text(_SPACE),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Italic(), new Text(_SPACE),
								new Text((request.TypeOfRest?.Name)))
							));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Bold(), new Text(_SPACE),
								new Text("Основание отказа в регистрации заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24), new Text(_SPACE),
								new Text((_FEDERAL_LAW)))
							));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24).Bold(), new Text(_SPACE),
								new Text("Причина отказа в регистрации заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_24), new Text(_SPACE),
								new Text((request.DeclineReason.Name)))
							));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					SignWorkerCompensationBlock(doc, account);

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Отказ" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}
		#endregion

		/// <summary>
		///     Уведомление о запрос необходимой информации
		/// </summary>
		public IDocument NotificationRequestInformation(Request request, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var isCompensationYouth = request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest;

					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ?? new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					if (isCompensationYouth)
					{
						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text("Уведомление о необходимости предоставления дополнительной информации для принятия решения о выплате компенсации"),
								new Break(),
								new Text("за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха"),
								new Break(),
								new Text("и оздоровления"))));
					}
					else
					{
						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
								new Text("Уведомление о необходимости предоставления дополнительной информации для принятия решения о выплате компенсации"),
								new Break(),
								new Text("за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления"))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									$"Уважаемый(ая) {applicant.LastName} {applicant.FirstName} {applicant.MiddleName}!"))));

					var t = isCompensationYouth ? "лицом из числа детей-сирот и детей, оставшихся без попечения родителей," : "родителями или иными законными представителями";

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text($"Для принятия решения о выплате компенсации за самостоятельно приобретенную {t} путевку для отдыха и оздоровления по Вашему заявлению от {request.DateRequest:dd.MM.yyyy}г. № {request.RequestNumber} Вам необходимо в течение 10 рабочих дней с момента направления данного уведомления на электронную почту, указанную в заявлении, предоставить в ГАУК \"МОСГОРТУР\" следующую информацию:"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									"..."))));

					if (isCompensationYouth)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("В случае непредоставления запрашиваемой информации в указанный срок, Вам будет отказано в осуществлении выплаты компенсации за самостоятельно приобретенную лицом из числа детей-сирот и детей, оставшихся без попечения родителей, путевку для отдыха и оздоровления."))));
					}
					else
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 },
									new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text("В случае непредоставления запрашиваемой информации в указанный срок, Вам будет отказано в осуществлении выплаты компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления."))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					SignWorkerCompensationBlock(doc, account);

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о необходимости предоставления дополнительной информации" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		///     Уведомление (базовая развилка)
		/// </summary>
		public IDocument NotificationDataSwitch(Request request, Account account)
		{
			var dr2 = ConfigurationManager.AppSettings["NotificationRefuseDeclineReasonWrongDocs"].LongParse() ?? 201904;
			var dr3 = ConfigurationManager.AppSettings["NotificationRefuseDeclineReasonQuota"].LongParse() ?? 201705;
			// отказ от вариантов
			var dr4 = ConfigurationManager.AppSettings["NotificationRefuseDeclineDiscardingOptions"].LongParse() ?? 201902;
			// неучастие в выборе альтернативного варианта (мини ЛОК 2020)
			var dr5 = ConfigurationManager.AppSettings["NotificationRefuseDeclineDiscardingChoose"].LongParse() ?? 201911;
			//неучастие заявителя во втором этапе заявочной кампании
			var dr6 = ConfigurationManager.AppSettings["NotParticipateInSecondStage"].LongParse() ?? 201911;

			if (request.StatusId == (long)StatusEnum.CancelByApplicant && !string.IsNullOrWhiteSpace(request.CertificateNumber))
			{
				var doc = NotificationDeadline(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return doc;
			}

			if (request.StatusId == (long)StatusEnum.CancelByApplicant)
			{
				return NotificationRefuse1090(request, account);
			}

			if (request.StatusId == (long)StatusEnum.Reject && request.DeclineReasonId == dr2)
			{
				var doc = NotificationRefuse10802(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return doc;
			}

			if (request.StatusId == (long)StatusEnum.Reject && request.DeclineReasonId == dr3)
			{
				var doc = NotificationRefuse10805(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return doc;
			}

			if (request.StatusId == (long)StatusEnum.Reject && request.DeclineReasonId == dr6)
			{
				var doc = NotificationRefuse108013(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return doc;
			}

			if (request.StatusId == (long)StatusEnum.Reject &&
				request.TypeOfRestId == (long)TypeOfRestEnum.Compensation)
			{
				var doc = NotificationRefuseCompensation(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return doc;
			}

			if (request.StatusId == (long)StatusEnum.Reject &&
				request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest)
			{
				var doc = NotificationRefuseCompensationYouthRest(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return doc;
			}

			if (request.StatusId == (long)StatusEnum.Reject)
			{
				var doc = NotificationRefuse1080(request, account);
				doc.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
				return doc;
			}

			var document = NotificationRefuseContent(request, account);
			document.RequestFileTypeId = (long)RequestFileTypeEnum.NotificationRefuse;
			return document;

		}

		/// <summary>
		///     Уведомление о регистрации заявления
		/// </summary>
		public IDocument AlternateCompanyNotification(IUnitOfWork unitOfWork, Account account, long requestId, string time = "")
		{
			var request = unitOfWork.GetById<Request>(requestId);
			var requests = unitOfWork.GetSet<Request>().Where(r => r.ParentRequestId == requestId).ToList();

			if (request.StatusId == (long)StatusEnum.DecisionMakingCovid)
				return AlternateCompanyRechoiseNotification(request, account, request.SourceId == (long)SourceEnum.Mpgu, time);

			return AlternateCompanyNotification(request, account, request.SourceId == (long)SourceEnum.Mpgu, request.StatusId == (long)StatusEnum.CancelByApplicant, time, requests);
		}

		/// <summary>
		///     получить данные по уведомлению
		/// </summary>
		public IDocument NotificationRefuseContentEx(Request request)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(),
							new Text("Уведомление об отказе в предоставлении услуг отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);

					var applicant = request.NullSafe(r => r.Applicant) ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{
									Space = SpaceProcessingModeValues.Preserve
								}),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"бесплатная путевка для отдыха и оздоровления/сертификат на получение выплаты на самостоятельную организацию отдыха и оздоровления"))));

					AppendChildrenToDocument(doc, request);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"отказ в предоставлении услуг отдыха и оздоровления в связи с отсутствием квоты на отдых и оздоровление."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Причина отказа: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"16 января 2019 г. проведено заседание комиссии в целях принятия решения о включении сведений о детях, находящихся в трудной жизненной ситуации, сопровождающих их лицах (в случае организации совместного выездного отдыха), лицах из числа детей сирот и детей, оставшихся без попечения родителей, в предварительный Реестр получателей услуг отдыха и оздоровления в соответствии с очередностью, которая учитывает предоставленные услуги отдыха и оздоровления в 2016-2018 годах, а также дату и время подачи заявления."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }, new Indentation { FirstLine = "500" }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"Установление очередности рассмотрения заявлений о предоставлении услуг отдыха и оздоровления осуществляется автоматизированной информационной системой \"Детский отдых\" в пределах утвержденного количества мест для соответствующего вида отдыха и оздоровления и соответствующей льготной категории детей."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }, new Indentation { FirstLine = "500" }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"В первую очередь рассмотрены заявления о предоставлении услуг отдыха и оздоровления детям, которым такие услуги в последние три года не предоставлялись."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }, new Indentation { FirstLine = "500" }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"В 2019 году детям, указанным в заявлениях, которым предоставлялись услуги отдыха и оздоровления в 2016-2018 годах, в связи с отсутствием мест в квоте на отдых и оздоровление для указанной льготной категории детей услуга не предоставляется и сведения о них не включены в предварительный Реестр получателей услуг отдыха и оздоровления."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\"."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отказе в предоставлении услуг отдыха и оздоровления" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		///     Список для ГИБДД
		/// </summary>
		public IDocument OrphanagePupilGroupListsGibddList(ListOfChilds list)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();

					Document document1 = new Document();

					Body body1 = new Body();

					Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "003B19CE", RsidParagraphAddition = "001C30EE", RsidParagraphProperties = "001C30EE", RsidRunAdditionDefault = "003B19CE", ParagraphId = "664A1426", TextId = "77777777" };

					ParagraphProperties paragraphProperties1 = new ParagraphProperties();
					SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
					Justification justification1 = new Justification() { Val = JustificationValues.Center };

					ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
					RunFonts runFonts1 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
					Bold bold1 = new Bold();
					BoldComplexScript boldComplexScript1 = new BoldComplexScript();
					FontSize fontSize1 = new FontSize() { Val = "28" };
					FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

					paragraphMarkRunProperties1.Append(runFonts1);
					paragraphMarkRunProperties1.Append(bold1);
					paragraphMarkRunProperties1.Append(boldComplexScript1);
					paragraphMarkRunProperties1.Append(fontSize1);
					paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

					paragraphProperties1.Append(spacingBetweenLines1);
					paragraphProperties1.Append(justification1);
					paragraphProperties1.Append(paragraphMarkRunProperties1);

					Run run1 = new Run() { RsidRunProperties = "003B19CE" };

					RunProperties runProperties1 = new RunProperties();
					RunFonts runFonts2 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
					Bold bold2 = new Bold();
					BoldComplexScript boldComplexScript2 = new BoldComplexScript();
					FontSize fontSize2 = new FontSize() { Val = "28" };
					FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

					runProperties1.Append(runFonts2);
					runProperties1.Append(bold2);
					runProperties1.Append(boldComplexScript2);
					runProperties1.Append(fontSize2);
					runProperties1.Append(fontSizeComplexScript2);
					Text text1 = new Text();
					text1.Text = "Информация для организации автомобильного трансфера";

					run1.Append(runProperties1);
					run1.Append(text1);

					paragraph1.Append(paragraphProperties1);
					paragraph1.Append(run1);

					Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "003B19CE", RsidParagraphAddition = "003B19CE", RsidParagraphProperties = "001C30EE", RsidRunAdditionDefault = "003B19CE", ParagraphId = "42F3D557", TextId = "77777777" };

					ParagraphProperties paragraphProperties2 = new ParagraphProperties();
					SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
					Justification justification2 = new Justification() { Val = JustificationValues.Center };

					ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
					RunFonts runFonts3 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
					Bold bold3 = new Bold();
					BoldComplexScript boldComplexScript3 = new BoldComplexScript();
					Color color1 = new Color() { Val = "000000", ThemeColor = ThemeColorValues.Text1 };
					FontSize fontSize3 = new FontSize() { Val = "28" };
					FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };

					paragraphMarkRunProperties2.Append(runFonts3);
					paragraphMarkRunProperties2.Append(bold3);
					paragraphMarkRunProperties2.Append(boldComplexScript3);
					paragraphMarkRunProperties2.Append(color1);
					paragraphMarkRunProperties2.Append(fontSize3);
					paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

					paragraphProperties2.Append(spacingBetweenLines2);
					paragraphProperties2.Append(justification2);
					paragraphProperties2.Append(paragraphMarkRunProperties2);

					paragraph2.Append(paragraphProperties2);

					//первая таблица
					Table table1 = new Table();

					TableProperties tableProperties1 = new TableProperties();
					TableStyle tableStyle1 = new TableStyle() { Val = "a3" };
					TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
					TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

					var tableBorders1 = new TableBorders(
						new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) });

					tableProperties1.Append(tableStyle1);
					tableProperties1.Append(tableWidth1);
					tableProperties1.Append(tableLook1);
					tableProperties1.Append(tableBorders1);

					TableGrid tableGrid1 = new TableGrid();
					GridColumn gridColumn1 = new GridColumn() { Width = "672" };
					GridColumn gridColumn2 = new GridColumn() { Width = "6203" };
					GridColumn gridColumn3 = new GridColumn() { Width = "2470" };

					tableGrid1.Append(gridColumn1);
					tableGrid1.Append(gridColumn2);
					tableGrid1.Append(gridColumn3);

					table1.Append(tableProperties1);
					table1.Append(tableGrid1);

					{
						TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "001C30EE", RsidTableRowAddition = "00917B74", RsidTableRowProperties = "00E20D38", ParagraphId = "353D4AF8", TextId = "77777777" };

						TableCell tableCell1 = new TableCell();

						TableCellProperties tableCellProperties1 = new TableCellProperties();
						TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "672", Type = TableWidthUnitValues.Dxa };
						TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

						tableCellProperties1.Append(tableCellWidth1);
						tableCellProperties1.Append(tableCellVerticalAlignment1);

						Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00FE30D3", RsidRunAdditionDefault = "00917B74", ParagraphId = "1F4C34C9", TextId = "77777777" };

						ParagraphProperties paragraphProperties3 = new ParagraphProperties();
						Justification justification3 = new Justification() { Val = JustificationValues.Center };

						ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
						RunFonts runFonts4 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize4 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

						paragraphMarkRunProperties3.Append(runFonts4);
						paragraphMarkRunProperties3.Append(fontSize4);
						paragraphMarkRunProperties3.Append(fontSizeComplexScript4);

						paragraphProperties3.Append(justification3);
						paragraphProperties3.Append(paragraphMarkRunProperties3);

						Run run2 = new Run() { RsidRunProperties = "00FE30D3" };

						RunProperties runProperties2 = new RunProperties();
						RunFonts runFonts5 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize5 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

						runProperties2.Append(runFonts5);
						runProperties2.Append(fontSize5);
						runProperties2.Append(fontSizeComplexScript5);
						Text text2 = new Text();
						text2.Text = "№ п/п/";

						run2.Append(runProperties2);
						run2.Append(text2);

						paragraph3.Append(paragraphProperties3);
						paragraph3.Append(run2);

						tableCell1.Append(tableCellProperties1);
						tableCell1.Append(paragraph3);

						TableCell tableCell2 = new TableCell();

						TableCellProperties tableCellProperties2 = new TableCellProperties();
						TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "6203", Type = TableWidthUnitValues.Dxa };
						TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

						tableCellProperties2.Append(tableCellWidth2);
						tableCellProperties2.Append(tableCellVerticalAlignment2);

						Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "008622BB", RsidRunAdditionDefault = "00917B74", ParagraphId = "7AD0013B", TextId = "63787869" };

						ParagraphProperties paragraphProperties4 = new ParagraphProperties();
						Justification justification4 = new Justification() { Val = JustificationValues.Center };

						ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
						RunFonts runFonts6 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize6 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

						paragraphMarkRunProperties4.Append(runFonts6);
						paragraphMarkRunProperties4.Append(fontSize6);
						paragraphMarkRunProperties4.Append(fontSizeComplexScript6);

						paragraphProperties4.Append(justification4);
						paragraphProperties4.Append(paragraphMarkRunProperties4);

						Run run3 = new Run() { RsidRunProperties = "00FE30D3" };

						RunProperties runProperties3 = new RunProperties();
						RunFonts runFonts7 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize7 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "28" };

						runProperties3.Append(runFonts7);
						runProperties3.Append(fontSize7);
						runProperties3.Append(fontSizeComplexScript7);
						Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
						text3.Text = "ФИО детей-сирот и воспитанников, помещенных ";

						run3.Append(runProperties3);
						run3.Append(text3);

						Run run4 = new Run() { RsidRunAddition = "00E67E42" };

						RunProperties runProperties4 = new RunProperties();
						RunFonts runFonts8 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize8 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "28" };

						runProperties4.Append(runFonts8);
						runProperties4.Append(fontSize8);
						runProperties4.Append(fontSizeComplexScript8);
						Break break1 = new Break();

						run4.Append(runProperties4);
						run4.Append(break1);

						Run run5 = new Run() { RsidRunProperties = "00FE30D3" };

						RunProperties runProperties5 = new RunProperties();
						RunFonts runFonts9 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize9 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "28" };

						runProperties5.Append(runFonts9);
						runProperties5.Append(fontSize9);
						runProperties5.Append(fontSizeComplexScript9);
						Text text4 = new Text() { Space = SpaceProcessingModeValues.Preserve };
						text4.Text = "в стационарные ";

						run5.Append(runProperties5);
						run5.Append(text4);

						Run run6 = new Run();

						RunProperties runProperties6 = new RunProperties();
						RunFonts runFonts10 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize10 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "28" };

						runProperties6.Append(runFonts10);
						runProperties6.Append(fontSize10);
						runProperties6.Append(fontSizeComplexScript10);
						Text text5 = new Text();
						text5.Text = "учреждения";

						run6.Append(runProperties6);
						run6.Append(text5);

						paragraph4.Append(paragraphProperties4);
						paragraph4.Append(run3);
						paragraph4.Append(run4);
						paragraph4.Append(run5);
						paragraph4.Append(run6);

						tableCell2.Append(tableCellProperties2);
						tableCell2.Append(paragraph4);

						TableCell tableCell3 = new TableCell();

						TableCellProperties tableCellProperties3 = new TableCellProperties();
						TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2470", Type = TableWidthUnitValues.Dxa };
						TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

						tableCellProperties3.Append(tableCellWidth3);
						tableCellProperties3.Append(tableCellVerticalAlignment3);

						Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "008622BB", RsidRunAdditionDefault = "00917B74", ParagraphId = "6B1239B8", TextId = "1FB04BCF" };

						ParagraphProperties paragraphProperties5 = new ParagraphProperties();
						Justification justification5 = new Justification() { Val = JustificationValues.Center };

						ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
						RunFonts runFonts11 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize11 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "28" };

						paragraphMarkRunProperties5.Append(runFonts11);
						paragraphMarkRunProperties5.Append(fontSize11);
						paragraphMarkRunProperties5.Append(fontSizeComplexScript11);

						paragraphProperties5.Append(justification5);
						paragraphProperties5.Append(paragraphMarkRunProperties5);

						Run run7 = new Run() { RsidRunProperties = "00FE30D3" };

						RunProperties runProperties7 = new RunProperties();
						RunFonts runFonts12 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize12 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "28" };

						runProperties7.Append(runFonts12);
						runProperties7.Append(fontSize12);
						runProperties7.Append(fontSizeComplexScript12);
						Text text6 = new Text();
						text6.Text = "Дата рождения";

						run7.Append(runProperties7);
						run7.Append(text6);

						paragraph5.Append(paragraphProperties5);
						paragraph5.Append(run7);

						tableCell3.Append(tableCellProperties3);
						tableCell3.Append(paragraph5);

						tableRow1.Append(tableCell1);
						tableRow1.Append(tableCell2);
						tableRow1.Append(tableCell3);

						table1.Append(tableRow1);
					}

					var gg = list.GroupPupils.Where(ss => ss.OrganisatonAddresId != null).ToList().GroupBy(g => g.OrganisatonAddresId);
					int i = 1;
					foreach (var a in gg)
					{
						{
							TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "001C30EE", RsidTableRowAddition = "00FE30D3", RsidTableRowProperties = "00E20D38", ParagraphId = "2AD6DBC6", TextId = "77777777" };

							TableCell tableCell4 = new TableCell();

							TableCellProperties tableCellProperties4 = new TableCellProperties();
							TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "9345", Type = TableWidthUnitValues.Dxa };
							GridSpan gridSpan1 = new GridSpan() { Val = 3 };
							TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

							tableCellProperties4.Append(tableCellWidth4);
							tableCellProperties4.Append(gridSpan1);
							tableCellProperties4.Append(tableCellVerticalAlignment4);

							Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00FE30D3", RsidParagraphProperties = "00FE30D3", RsidRunAdditionDefault = "00FE30D3", ParagraphId = "0B71E124", TextId = "47E8F151" };

							ParagraphProperties paragraphProperties6 = new ParagraphProperties();

							ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
							RunFonts runFonts13 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							Bold bold4 = new Bold();
							FontSize fontSize13 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "28" };

							paragraphMarkRunProperties6.Append(runFonts13);
							paragraphMarkRunProperties6.Append(bold4);
							paragraphMarkRunProperties6.Append(fontSize13);
							paragraphMarkRunProperties6.Append(fontSizeComplexScript13);

							paragraphProperties6.Append(paragraphMarkRunProperties6);

							Run run8 = new Run() { RsidRunProperties = "00FE30D3" };

							RunProperties runProperties8 = new RunProperties();
							RunFonts runFonts14 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							Bold bold5 = new Bold();
							FontSize fontSize14 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "28" };

							runProperties8.Append(runFonts14);
							runProperties8.Append(bold5);
							runProperties8.Append(fontSize14);
							runProperties8.Append(fontSizeComplexScript14);
							Text text7 = new Text();
							text7.Text = $"Адрес № {i++}: ";

							run8.Append(runProperties8);
							run8.Append(text7);

							Run run9 = new Run() { RsidRunAddition = "00B77DFD" };

							RunProperties runProperties9 = new RunProperties();
							RunFonts runFonts15 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							Bold bold6 = new Bold();
							FontSize fontSize15 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "28" };

							runProperties9.Append(runFonts15);
							runProperties9.Append(bold6);
							runProperties9.Append(fontSize15);
							runProperties9.Append(fontSizeComplexScript15);
							Text text8 = new Text() { Space = SpaceProcessingModeValues.Preserve };
							text8.Text = $" {a.FirstOrDefault().OrganisatonAddres.Address.Name}";

							run9.Append(runProperties9);
							run9.Append(text8);

							paragraph6.Append(paragraphProperties6);
							paragraph6.Append(run8);
							paragraph6.Append(run9);

							tableCell4.Append(tableCellProperties4);
							tableCell4.Append(paragraph6);

							tableRow2.Append(tableCell4);

							table1.Append(tableRow2);
						}

						int j = 1;
						foreach (var p in a.ToList())
						{
							TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "001C30EE", RsidTableRowAddition = "00917B74", RsidTableRowProperties = "00B830E4", ParagraphId = "788A9D53", TextId = "77777777" };

							TableCell tableCell5 = new TableCell();

							TableCellProperties tableCellProperties5 = new TableCellProperties();
							TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "672", Type = TableWidthUnitValues.Dxa };
							TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

							tableCellProperties5.Append(tableCellWidth5);
							tableCellProperties5.Append(tableCellVerticalAlignment5);

							Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00FE30D3", RsidRunAdditionDefault = "00917B74", ParagraphId = "0FBA0F87", TextId = "77777777" };

							ParagraphProperties paragraphProperties7 = new ParagraphProperties();

							ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
							RunFonts runFonts16 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize16 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "28" };

							paragraphMarkRunProperties7.Append(runFonts16);
							paragraphMarkRunProperties7.Append(fontSize16);
							paragraphMarkRunProperties7.Append(fontSizeComplexScript16);

							paragraphProperties7.Append(paragraphMarkRunProperties7);

							Run run10 = new Run() { RsidRunProperties = "00FE30D3" };

							RunProperties runProperties10 = new RunProperties();
							RunFonts runFonts17 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize17 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "28" };

							runProperties10.Append(runFonts17);
							runProperties10.Append(fontSize17);
							runProperties10.Append(fontSizeComplexScript17);
							Text text9 = new Text();
							text9.Text = $"{j++}.";

							run10.Append(runProperties10);
							run10.Append(text9);

							paragraph7.Append(paragraphProperties7);
							paragraph7.Append(run10);

							tableCell5.Append(tableCellProperties5);
							tableCell5.Append(paragraph7);

							TableCell tableCell6 = new TableCell();

							TableCellProperties tableCellProperties6 = new TableCellProperties();
							TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "6203", Type = TableWidthUnitValues.Dxa };
							TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

							tableCellProperties6.Append(tableCellWidth6);
							tableCellProperties6.Append(tableCellVerticalAlignment6);

							Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "005A7013", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00FE30D3", RsidRunAdditionDefault = "005A7013", ParagraphId = "6FF2644A", TextId = "61911ABD" };

							ParagraphProperties paragraphProperties8 = new ParagraphProperties();

							ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
							RunFonts runFonts18 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize18 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "28" };

							paragraphMarkRunProperties8.Append(runFonts18);
							paragraphMarkRunProperties8.Append(fontSize18);
							paragraphMarkRunProperties8.Append(fontSizeComplexScript18);

							paragraphProperties8.Append(paragraphMarkRunProperties8);
							ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

							Run run11 = new Run();

							RunProperties runProperties11 = new RunProperties();
							RunFonts runFonts19 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize19 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "28" };

							runProperties11.Append(runFonts19);
							runProperties11.Append(fontSize19);
							runProperties11.Append(fontSizeComplexScript19);
							Text text10 = new Text();
							text10.Text = p.Pupil.Child.GetFio();

							run11.Append(runProperties11);
							run11.Append(text10);

							paragraph8.Append(paragraphProperties8);
							paragraph8.Append(proofError1);
							paragraph8.Append(run11);

							tableCell6.Append(tableCellProperties6);
							tableCell6.Append(paragraph8);

							TableCell tableCell7 = new TableCell();

							TableCellProperties tableCellProperties7 = new TableCellProperties();
							TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "2470", Type = TableWidthUnitValues.Dxa };
							TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

							tableCellProperties7.Append(tableCellWidth7);
							tableCellProperties7.Append(tableCellVerticalAlignment7);

							Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00FE30D3", RsidRunAdditionDefault = "005A7013", ParagraphId = "40FE5834", TextId = "46CBE13F" };

							ParagraphProperties paragraphProperties9 = new ParagraphProperties();

							ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
							RunFonts runFonts21 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize21 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "28" };

							paragraphMarkRunProperties9.Append(runFonts21);
							paragraphMarkRunProperties9.Append(fontSize21);
							paragraphMarkRunProperties9.Append(fontSizeComplexScript21);

							paragraphProperties9.Append(paragraphMarkRunProperties9);

							Run run13 = new Run();

							RunProperties runProperties13 = new RunProperties();
							RunFonts runFonts22 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize22 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "28" };

							runProperties13.Append(runFonts22);
							runProperties13.Append(fontSize22);
							runProperties13.Append(fontSizeComplexScript22);
							Text text12 = new Text();
							text12.Text = $"{p.Pupil.Child.DateOfBirth.FormatEx(string.Empty)}";

							run13.Append(runProperties13);
							run13.Append(text12);

							paragraph9.Append(paragraphProperties9);
							paragraph9.Append(run13);

							tableCell7.Append(tableCellProperties7);
							tableCell7.Append(paragraph9);

							tableRow3.Append(tableCell5);
							tableRow3.Append(tableCell6);
							tableRow3.Append(tableCell7);

							table1.Append(tableRow3);
						}
					}

					Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "001C30EE", RsidParagraphProperties = "00FE30D3", RsidRunAdditionDefault = "001C30EE", ParagraphId = "19850A2D", TextId = "77777777" };

					ParagraphProperties paragraphProperties10 = new ParagraphProperties();

					ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
					RunFonts runFonts23 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
					FontSize fontSize23 = new FontSize() { Val = "24" };
					FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "24" };

					paragraphMarkRunProperties10.Append(runFonts23);
					paragraphMarkRunProperties10.Append(fontSize23);
					paragraphMarkRunProperties10.Append(fontSizeComplexScript23);

					paragraphProperties10.Append(paragraphMarkRunProperties10);

					paragraph10.Append(paragraphProperties10);



					//вторая таблица
					Table table2 = new Table();

					TableProperties tableProperties2 = new TableProperties();
					TableStyle tableStyle2 = new TableStyle() { Val = "a3" };
					TableWidth tableWidth2 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
					TableLook tableLook2 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };
					var tableBorders2 = new TableBorders(
						new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
						new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) });

					tableProperties2.Append(tableStyle2);
					tableProperties2.Append(tableWidth2);
					tableProperties2.Append(tableLook2);
					tableProperties2.Append(tableBorders2);

					TableGrid tableGrid2 = new TableGrid();
					GridColumn gridColumn4 = new GridColumn() { Width = "676" };
					GridColumn gridColumn5 = new GridColumn() { Width = "4793" };
					GridColumn gridColumn6 = new GridColumn() { Width = "1547" };
					GridColumn gridColumn7 = new GridColumn() { Width = "2329" };

					tableGrid2.Append(gridColumn4);
					tableGrid2.Append(gridColumn5);
					tableGrid2.Append(gridColumn6);
					tableGrid2.Append(gridColumn7);

					table2.Append(tableProperties2);
					table2.Append(tableGrid2);

					{
						TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "001C30EE", RsidTableRowAddition = "00917B74", RsidTableRowProperties = "00E20D38", ParagraphId = "6ABEDCDC", TextId = "77777777" };

						TableCell tableCell8 = new TableCell();

						TableCellProperties tableCellProperties8 = new TableCellProperties();
						TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "676", Type = TableWidthUnitValues.Dxa };
						TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

						tableCellProperties8.Append(tableCellWidth8);
						tableCellProperties8.Append(tableCellVerticalAlignment8);

						Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00821F51", RsidRunAdditionDefault = "00917B74", ParagraphId = "1FE426A2", TextId = "77777777" };

						ParagraphProperties paragraphProperties11 = new ParagraphProperties();
						Justification justification6 = new Justification() { Val = JustificationValues.Center };

						ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
						RunFonts runFonts24 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize24 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "28" };

						paragraphMarkRunProperties11.Append(runFonts24);
						paragraphMarkRunProperties11.Append(fontSize24);
						paragraphMarkRunProperties11.Append(fontSizeComplexScript24);

						paragraphProperties11.Append(justification6);
						paragraphProperties11.Append(paragraphMarkRunProperties11);

						Run run14 = new Run() { RsidRunProperties = "00FE30D3" };

						RunProperties runProperties14 = new RunProperties();
						RunFonts runFonts25 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize25 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "28" };

						runProperties14.Append(runFonts25);
						runProperties14.Append(fontSize25);
						runProperties14.Append(fontSizeComplexScript25);
						Text text13 = new Text();
						text13.Text = "№ п/п/";

						run14.Append(runProperties14);
						run14.Append(text13);

						paragraph11.Append(paragraphProperties11);
						paragraph11.Append(run14);

						tableCell8.Append(tableCellProperties8);
						tableCell8.Append(paragraph11);

						TableCell tableCell9 = new TableCell();

						TableCellProperties tableCellProperties9 = new TableCellProperties();
						TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "4793", Type = TableWidthUnitValues.Dxa };
						TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

						tableCellProperties9.Append(tableCellWidth9);
						tableCellProperties9.Append(tableCellVerticalAlignment9);

						Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00E20D38", RsidRunAdditionDefault = "00917B74", ParagraphId = "4B01B719", TextId = "61855BE1" };

						ParagraphProperties paragraphProperties12 = new ParagraphProperties();
						Justification justification7 = new Justification() { Val = JustificationValues.Center };

						ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
						RunFonts runFonts26 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize26 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "28" };

						paragraphMarkRunProperties12.Append(runFonts26);
						paragraphMarkRunProperties12.Append(fontSize26);
						paragraphMarkRunProperties12.Append(fontSizeComplexScript26);

						paragraphProperties12.Append(justification7);
						paragraphProperties12.Append(paragraphMarkRunProperties12);

						Run run15 = new Run() { RsidRunProperties = "00FE30D3" };

						RunProperties runProperties15 = new RunProperties();
						RunFonts runFonts27 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize27 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "28" };

						runProperties15.Append(runFonts27);
						runProperties15.Append(fontSize27);
						runProperties15.Append(fontSizeComplexScript27);
						Text text14 = new Text() { Space = SpaceProcessingModeValues.Preserve };
						text14.Text = "ФИО ";

						run15.Append(runProperties15);
						run15.Append(text14);

						Run run16 = new Run();

						RunProperties runProperties16 = new RunProperties();
						RunFonts runFonts28 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize28 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "28" };

						runProperties16.Append(runFonts28);
						runProperties16.Append(fontSize28);
						runProperties16.Append(fontSizeComplexScript28);
						Text text15 = new Text();
						text15.Text = "Сопровождающих";

						run16.Append(runProperties16);
						run16.Append(text15);

						paragraph12.Append(paragraphProperties12);
						paragraph12.Append(run15);
						paragraph12.Append(run16);

						tableCell9.Append(tableCellProperties9);
						tableCell9.Append(paragraph12);

						TableCell tableCell10 = new TableCell();

						TableCellProperties tableCellProperties10 = new TableCellProperties();
						TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "1547", Type = TableWidthUnitValues.Dxa };
						TableCellVerticalAlignment tableCellVerticalAlignment10 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

						tableCellProperties10.Append(tableCellWidth10);
						tableCellProperties10.Append(tableCellVerticalAlignment10);

						Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00821F51", RsidRunAdditionDefault = "00917B74", ParagraphId = "1BDA199E", TextId = "77777777" };

						ParagraphProperties paragraphProperties13 = new ParagraphProperties();
						Justification justification8 = new Justification() { Val = JustificationValues.Center };

						ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
						RunFonts runFonts29 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize29 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "28" };

						paragraphMarkRunProperties13.Append(runFonts29);
						paragraphMarkRunProperties13.Append(fontSize29);
						paragraphMarkRunProperties13.Append(fontSizeComplexScript29);

						paragraphProperties13.Append(justification8);
						paragraphProperties13.Append(paragraphMarkRunProperties13);

						Run run17 = new Run() { RsidRunProperties = "00FE30D3" };

						RunProperties runProperties17 = new RunProperties();
						RunFonts runFonts30 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize30 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "28" };

						runProperties17.Append(runFonts30);
						runProperties17.Append(fontSize30);
						runProperties17.Append(fontSizeComplexScript30);
						Text text16 = new Text();
						text16.Text = "Дата рождения";

						run17.Append(runProperties17);
						run17.Append(text16);

						paragraph13.Append(paragraphProperties13);
						paragraph13.Append(run17);

						tableCell10.Append(tableCellProperties10);
						tableCell10.Append(paragraph13);

						TableCell tableCell11 = new TableCell();

						TableCellProperties tableCellProperties11 = new TableCellProperties();
						TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "2329", Type = TableWidthUnitValues.Dxa };
						TableCellVerticalAlignment tableCellVerticalAlignment11 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

						tableCellProperties11.Append(tableCellWidth11);
						tableCellProperties11.Append(tableCellVerticalAlignment11);

						Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00821F51", RsidRunAdditionDefault = "00917B74", ParagraphId = "19158C95", TextId = "77777777" };

						ParagraphProperties paragraphProperties14 = new ParagraphProperties();
						Justification justification9 = new Justification() { Val = JustificationValues.Center };

						ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
						RunFonts runFonts31 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize31 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "28" };

						paragraphMarkRunProperties14.Append(runFonts31);
						paragraphMarkRunProperties14.Append(fontSize31);
						paragraphMarkRunProperties14.Append(fontSizeComplexScript31);

						paragraphProperties14.Append(justification9);
						paragraphProperties14.Append(paragraphMarkRunProperties14);

						Run run18 = new Run();

						RunProperties runProperties18 = new RunProperties();
						RunFonts runFonts32 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
						FontSize fontSize32 = new FontSize() { Val = "28" };
						FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "28" };

						runProperties18.Append(runFonts32);
						runProperties18.Append(fontSize32);
						runProperties18.Append(fontSizeComplexScript32);
						Text text17 = new Text();
						text17.Text = "Контактный телефон";

						run18.Append(runProperties18);
						run18.Append(text17);

						paragraph14.Append(paragraphProperties14);
						paragraph14.Append(run18);

						tableCell11.Append(tableCellProperties11);
						tableCell11.Append(paragraph14);

						tableRow4.Append(tableCell8);
						tableRow4.Append(tableCell9);
						tableRow4.Append(tableCell10);
						tableRow4.Append(tableCell11);

						table2.Append(tableRow4);
					}

					var gh = list.GroupCollaborators.Where(ss => ss.OrganisatonAddresId != null).ToList().GroupBy(g => g.OrganisatonAddresId);
					i = 1;
					foreach (var a in gh)
					{
						{
							TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "001C30EE", RsidTableRowAddition = "00917B74", RsidTableRowProperties = "00E20D38", ParagraphId = "26A2766A", TextId = "77777777" };

							TableCell tableCell12 = new TableCell();

							TableCellProperties tableCellProperties12 = new TableCellProperties();
							TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "9345", Type = TableWidthUnitValues.Dxa };
							GridSpan gridSpan2 = new GridSpan() { Val = 4 };
							TableCellVerticalAlignment tableCellVerticalAlignment12 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

							tableCellProperties12.Append(tableCellWidth12);
							tableCellProperties12.Append(gridSpan2);
							tableCellProperties12.Append(tableCellVerticalAlignment12);

							Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00821F51", RsidRunAdditionDefault = "00917B74", ParagraphId = "372BCE9F", TextId = "681DBDC7" };

							ParagraphProperties paragraphProperties15 = new ParagraphProperties();

							ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
							RunFonts runFonts33 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							Bold bold7 = new Bold();
							FontSize fontSize33 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "28" };

							paragraphMarkRunProperties15.Append(runFonts33);
							paragraphMarkRunProperties15.Append(bold7);
							paragraphMarkRunProperties15.Append(fontSize33);
							paragraphMarkRunProperties15.Append(fontSizeComplexScript33);

							paragraphProperties15.Append(paragraphMarkRunProperties15);

							Run run19 = new Run() { RsidRunProperties = "00FE30D3" };

							RunProperties runProperties19 = new RunProperties();
							RunFonts runFonts34 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							Bold bold8 = new Bold();
							FontSize fontSize34 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "28" };

							runProperties19.Append(runFonts34);
							runProperties19.Append(bold8);
							runProperties19.Append(fontSize34);
							runProperties19.Append(fontSizeComplexScript34);
							Text text18 = new Text();
							text18.Text = $"Адрес № {i++}: ";

							run19.Append(runProperties19);
							run19.Append(text18);

							Run run20 = new Run();

							RunProperties runProperties20 = new RunProperties();
							RunFonts runFonts35 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							Bold bold9 = new Bold();
							FontSize fontSize35 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "28" };

							runProperties20.Append(runFonts35);
							runProperties20.Append(bold9);
							runProperties20.Append(fontSize35);
							runProperties20.Append(fontSizeComplexScript35);
							Text text19 = new Text() { Space = SpaceProcessingModeValues.Preserve };
							text19.Text = $" {a.FirstOrDefault().OrganisatonAddres.Address.Name}";

							run20.Append(runProperties20);
							run20.Append(text19);

							paragraph15.Append(paragraphProperties15);
							paragraph15.Append(run19);
							paragraph15.Append(run20);

							tableCell12.Append(tableCellProperties12);
							tableCell12.Append(paragraph15);

							tableRow5.Append(tableCell12);

							table2.Append(tableRow5);
						}

						var j = 1;
						foreach (var p in a.ToList())
						{
							TableRow tableRow6 = new TableRow() { RsidTableRowMarkRevision = "001C30EE", RsidTableRowAddition = "00917B74", RsidTableRowProperties = "00E20D38", ParagraphId = "40D2E31E", TextId = "77777777" };

							TableCell tableCell13 = new TableCell();

							TableCellProperties tableCellProperties13 = new TableCellProperties();
							TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "676", Type = TableWidthUnitValues.Dxa };
							TableCellVerticalAlignment tableCellVerticalAlignment13 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

							tableCellProperties13.Append(tableCellWidth13);
							tableCellProperties13.Append(tableCellVerticalAlignment13);

							Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00821F51", RsidRunAdditionDefault = "00917B74", ParagraphId = "41A59361", TextId = "77777777" };

							ParagraphProperties paragraphProperties16 = new ParagraphProperties();

							ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
							RunFonts runFonts36 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize36 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "28" };

							paragraphMarkRunProperties16.Append(runFonts36);
							paragraphMarkRunProperties16.Append(fontSize36);
							paragraphMarkRunProperties16.Append(fontSizeComplexScript36);

							paragraphProperties16.Append(paragraphMarkRunProperties16);

							Run run21 = new Run() { RsidRunProperties = "00FE30D3" };

							RunProperties runProperties21 = new RunProperties();
							RunFonts runFonts37 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize37 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "28" };

							runProperties21.Append(runFonts37);
							runProperties21.Append(fontSize37);
							runProperties21.Append(fontSizeComplexScript37);
							Text text20 = new Text();
							text20.Text = $"{j++}.";

							run21.Append(runProperties21);
							run21.Append(text20);

							paragraph16.Append(paragraphProperties16);
							paragraph16.Append(run21);

							tableCell13.Append(tableCellProperties13);
							tableCell13.Append(paragraph16);

							TableCell tableCell14 = new TableCell();

							TableCellProperties tableCellProperties14 = new TableCellProperties();
							TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "4793", Type = TableWidthUnitValues.Dxa };
							TableCellVerticalAlignment tableCellVerticalAlignment14 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

							tableCellProperties14.Append(tableCellWidth14);
							tableCellProperties14.Append(tableCellVerticalAlignment14);

							Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00821F51", RsidRunAdditionDefault = "005A7013", ParagraphId = "0813C609", TextId = "5DFE300D" };

							ParagraphProperties paragraphProperties17 = new ParagraphProperties();

							ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
							RunFonts runFonts38 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize38 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "28" };

							paragraphMarkRunProperties17.Append(runFonts38);
							paragraphMarkRunProperties17.Append(fontSize38);
							paragraphMarkRunProperties17.Append(fontSizeComplexScript38);

							paragraphProperties17.Append(paragraphMarkRunProperties17);
							ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.SpellStart };

							Run run22 = new Run();

							RunProperties runProperties22 = new RunProperties();
							RunFonts runFonts39 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize39 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "28" };

							runProperties22.Append(runFonts39);
							runProperties22.Append(fontSize39);
							runProperties22.Append(fontSizeComplexScript39);
							Text text21 = new Text();
							text21.Text = p.OrganisatonCollaborator.Applicant.GetFio();

							run22.Append(runProperties22);
							run22.Append(text21);

							paragraph17.Append(paragraphProperties17);
							paragraph17.Append(proofError3);
							paragraph17.Append(run22);

							tableCell14.Append(tableCellProperties14);
							tableCell14.Append(paragraph17);

							TableCell tableCell15 = new TableCell();

							TableCellProperties tableCellProperties15 = new TableCellProperties();
							TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "1547", Type = TableWidthUnitValues.Dxa };
							TableCellVerticalAlignment tableCellVerticalAlignment15 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

							tableCellProperties15.Append(tableCellWidth15);
							tableCellProperties15.Append(tableCellVerticalAlignment15);

							Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00821F51", RsidRunAdditionDefault = "005A7013", ParagraphId = "79A0F9EB", TextId = "7DF84B01" };

							ParagraphProperties paragraphProperties18 = new ParagraphProperties();

							ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
							RunFonts runFonts41 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize41 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "28" };

							paragraphMarkRunProperties18.Append(runFonts41);
							paragraphMarkRunProperties18.Append(fontSize41);
							paragraphMarkRunProperties18.Append(fontSizeComplexScript41);

							paragraphProperties18.Append(paragraphMarkRunProperties18);

							Run run24 = new Run();

							RunProperties runProperties24 = new RunProperties();
							RunFonts runFonts42 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize42 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "28" };

							runProperties24.Append(runFonts42);
							runProperties24.Append(fontSize42);
							runProperties24.Append(fontSizeComplexScript42);
							Text text23 = new Text();
							text23.Text = p.OrganisatonCollaborator.Applicant.DateOfBirth.FormatEx(string.Empty);

							run24.Append(runProperties24);
							run24.Append(text23);

							paragraph18.Append(paragraphProperties18);
							paragraph18.Append(run24);

							tableCell15.Append(tableCellProperties15);
							tableCell15.Append(paragraph18);

							TableCell tableCell16 = new TableCell();

							TableCellProperties tableCellProperties16 = new TableCellProperties();
							TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "2329", Type = TableWidthUnitValues.Dxa };
							TableCellVerticalAlignment tableCellVerticalAlignment16 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

							tableCellProperties16.Append(tableCellWidth16);
							tableCellProperties16.Append(tableCellVerticalAlignment16);

							Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00FE30D3", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00821F51", RsidRunAdditionDefault = "005A7013", ParagraphId = "2FA0A47C", TextId = "137AC9D2" };

							ParagraphProperties paragraphProperties19 = new ParagraphProperties();

							ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
							RunFonts runFonts43 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize43 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "28" };

							paragraphMarkRunProperties19.Append(runFonts43);
							paragraphMarkRunProperties19.Append(fontSize43);
							paragraphMarkRunProperties19.Append(fontSizeComplexScript43);

							paragraphProperties19.Append(paragraphMarkRunProperties19);

							Run run25 = new Run();

							RunProperties runProperties25 = new RunProperties();
							RunFonts runFonts44 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
							FontSize fontSize44 = new FontSize() { Val = "28" };
							FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "28" };

							runProperties25.Append(runFonts44);
							runProperties25.Append(fontSize44);
							runProperties25.Append(fontSizeComplexScript44);
							Text text24 = new Text();
							text24.Text = p.OrganisatonCollaborator.Applicant.Phone;

							run25.Append(runProperties25);
							run25.Append(text24);

							paragraph19.Append(paragraphProperties19);
							paragraph19.Append(run25);

							tableCell16.Append(tableCellProperties16);
							tableCell16.Append(paragraph19);

							tableRow6.Append(tableCell13);
							tableRow6.Append(tableCell14);
							tableRow6.Append(tableCell15);
							tableRow6.Append(tableCell16);

							table2.Append(tableRow6);
						}
					}



					Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "001C30EE", RsidParagraphAddition = "00917B74", RsidParagraphProperties = "00DD566E", RsidRunAdditionDefault = "00917B74", ParagraphId = "4EC3A709", TextId = "77777777" };

					ParagraphProperties paragraphProperties20 = new ParagraphProperties();

					ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
					RunFonts runFonts45 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
					FontSize fontSize45 = new FontSize() { Val = "24" };
					FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "24" };

					paragraphMarkRunProperties20.Append(runFonts45);
					paragraphMarkRunProperties20.Append(fontSize45);
					paragraphMarkRunProperties20.Append(fontSizeComplexScript45);

					paragraphProperties20.Append(paragraphMarkRunProperties20);

					paragraph20.Append(paragraphProperties20);

					SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "001C30EE", RsidR = "00917B74", RsidSect = "000C0FCD" };
					PageSize pageSize1 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
					PageMargin pageMargin1 = new PageMargin() { Top = 1134, Right = (UInt32Value)850U, Bottom = 1134, Left = (UInt32Value)1701U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U };
					Columns columns1 = new Columns() { Space = "708" };
					DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

					sectionProperties1.Append(pageSize1);
					sectionProperties1.Append(pageMargin1);
					sectionProperties1.Append(columns1);
					sectionProperties1.Append(docGrid1);

					body1.Append(paragraph1);
					body1.Append(paragraph2);
					body1.Append(table1);
					body1.Append(paragraph10);
					body1.Append(table2);
					body1.Append(paragraph20);
					body1.Append(sectionProperties1);

					document1.Append(body1);

					mainPart.Document = document1;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Список для ГИБДД" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		///     Перечень информации о детях
		/// </summary>
		public IDocument TradeUnionWord(IUnitOfWork unitOfWork, TradeUnionWordFilter filter)
		{
			var list = unitOfWork.GetById<TradeUnionList>(filter.TradeUnionId);


			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var body = new Body();
					var sectionPropertys = new SectionProperties();
					sectionPropertys.AppendChild(new PageSize
					{ Orient = PageOrientationValues.Landscape, Width = 15840, Height = 12240 });
					sectionPropertys.AppendChild(new PageMargin
					{
						Top = 851,
						Right = 567U,
						Bottom = 567,
						Left = 567U,
						Header = 720U,
						Footer = 720U,
						Gutter = 0U
					});
					body.AppendChild(sectionPropertys);
					var doc = new Document(body);

					mainPart.Document = doc;
					var titleProp = new RunProperties().SetFont().SetFontSize("28");
					var titlePropBold = new RunProperties().SetFont().SetFontSize("28").Bold();
					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(titleProp.CloneNode(true),
							new Text(
								"Перечень информации о детях, принявших участие в оздоровительных мероприятиях в организациях отдыха и оздоровления, с выделением субсидии из бюджета города Москвы"),
							new Break()),
						new Run(titlePropBold.CloneNode(true),
							new Text($"{list?.Camp?.Name}"),
							new Break()),
						new Run(titlePropBold.CloneNode(true),
							new Text(
								$"{list?.GroupedTimeOfRest?.Name} {list?.YearOfRest?.Year} года (с {list?.DateFrom.FormatEx()} по {list?.DateTo.FormatEx()})")
						)));

					var table = new Table();

					var tblProp = new TableProperties();

					var tableMainStyle = new TableStyle { Val = "TableGrid" };
					var tableMainWidth = new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
					tblProp.Append(tableMainStyle, tableMainWidth);

					table.AppendChild(tblProp);

					var tg = new TableGrid(
						new GridColumn { Width = "188" },
						new GridColumn { Width = "616" },
						new GridColumn { Width = "476" },
						new GridColumn { Width = "1041" },
						new GridColumn { Width = "851" },
						new GridColumn { Width = "616" },
						new GridColumn { Width = "571" },
						new GridColumn { Width = "616" }
					);
					table.AppendChild(tg);

					var tr = new TableRow();
					tr.Text("№",
						new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
							.Width("188"),
						new ParagraphProperties().CenterAlign().NoSpacing(),
						new RunProperties().SetFont().SetFontSize("20"));
					tr.Text("Ф.И.О.",
						new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
							.Width("616"),
						new ParagraphProperties().CenterAlign().NoSpacing(),
						new RunProperties().SetFont().SetFontSize("20"));
					tr.Text("Дата рождения",
						new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
							.Width("476"),
						new ParagraphProperties().CenterAlign().NoSpacing(),
						new RunProperties().SetFont().SetFontSize("20"));
					tr.Text("Наименование и реквизиты документа, удостоверяющего личность",
						new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
							.Width("1041"),
						new ParagraphProperties().CenterAlign().NoSpacing(),
						new RunProperties().SetFont().SetFontSize("20"));
					tr.Text("Место жительства",
						new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
							.Width("851"),
						new ParagraphProperties().CenterAlign().NoSpacing(),
						new RunProperties().SetFont().SetFontSize("20"));
					tr.Text("ФИО законного представителя несовершеннолетнего лица",
						new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
							.Width("616"),
						new ParagraphProperties().CenterAlign().NoSpacing(),
						new RunProperties().SetFont().SetFontSize("20"));
					tr.Text("Контактный телефон законного представителя несовершеннолетнего лица",
						new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
							.Width("571"),
						new ParagraphProperties().CenterAlign().NoSpacing(),
						new RunProperties().SetFont().SetFontSize("20"));
					tr.Text("Место работы родителя",
						new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
							.Width("616"),
						new ParagraphProperties().CenterAlign().NoSpacing(),
						new RunProperties().SetFont().SetFontSize("20"));
					table.AppendChild(tr);

					var children = list?.Campers?.Where(ss => !(filter.CameChildren ?? false) || ss.IsChecked).OrderBy(c => c.Child?.LastName).ThenBy(c => c.Child?.FirstName).ToList() ?? new List<TradeUnionCamper>();

					var index = 1;
					foreach (var child in children)
					{
						tr = new TableRow();
						tr.Text(index.ToString(),
							new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
								.Width("188"),
							new ParagraphProperties().CenterAlign().NoSpacing(),
							new RunProperties().SetFont().SetFontSize("20"));
						tr.Text($"{child?.Child?.LastName} {child?.Child?.FirstName} {child?.Child?.MiddleName}",
							new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
								.Width("616"),
							new ParagraphProperties().CenterAlign().NoSpacing(),
							new RunProperties().SetFont().SetFontSize("20"));
						tr.Text($"{child?.Child?.DateOfBirth.FormatEx()}",
							new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
								.Width("476"),
							new ParagraphProperties().CenterAlign().NoSpacing(),
							new RunProperties().SetFont().SetFontSize("20"));
						tr.Text(
							$"{child?.Child?.DocumentType?.Name} {child?.Child?.DocumentSeria} №{child?.Child?.DocumentNumber}, выдано {child?.Child?.DocumentSubjectIssue} {child?.Child?.DocumentDateOfIssue.FormatEx()}",
							new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
								.Width("1041"),
							new ParagraphProperties().CenterAlign().NoSpacing(),
							new RunProperties().SetFont().SetFontSize("20"));
						tr.Text(child?.Child?.Address?.ToString() ?? child?.AddressChild,
							new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
								.Width("851"),
							new ParagraphProperties().CenterAlign().NoSpacing(),
							new RunProperties().SetFont().SetFontSize("20"));
						tr.Text($"{child?.Parent?.LastName} {child?.Parent?.FirstName} {child?.Parent?.MiddleName}",
							new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
								.Width("616"),
							new ParagraphProperties().CenterAlign().NoSpacing(),
							new RunProperties().SetFont().SetFontSize("20"));
						tr.Text($"{child?.Parent?.Phone}",
							new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
								.Width("571"),
							new ParagraphProperties().CenterAlign().NoSpacing(),
							new RunProperties().SetFont().SetFontSize("20"));
						tr.Text($"{child?.ParentPlaceWork}",
							new TableCellProperties().Borders(new TableCellBorders().AllBorder()).CenterVAlign()
								.Width("616"),
							new ParagraphProperties().CenterAlign().NoSpacing(),
							new RunProperties().SetFont().SetFontSize("20"));
						table.AppendChild(tr);
						index++;
					}

					doc.AppendChild(table);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_16).Bold(), new Text(_SPACE)))); doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_16).Bold(), new Text(_SPACE))));


					//блок подписей
					var table2 = new Table();

					var tableProperties2 = new TableProperties();
					var tableWidth2 = new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto };
					var tableLook2 = new TableLook { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

					tableProperties2.Append(tableWidth2);
					tableProperties2.Append(tableLook2);

					var tableGrid2 = new TableGrid();
					var gridColumn9 = new GridColumn { Width = "3417" };
					var gridColumn10 = new GridColumn { Width = "3417" };
					var gridColumn11 = new GridColumn { Width = "2916" };

					tableGrid2.Append(gridColumn9);
					tableGrid2.Append(gridColumn10);
					tableGrid2.Append(gridColumn11);

					var tableRow5 = new TableRow { RsidTableRowMarkRevision = "00232403", RsidTableRowAddition = "004950F8", RsidTableRowProperties = "00751945", ParagraphId = "68AE8A96", TextId = "77777777" };

					var tableCell33 = new TableCell();

					var tableCellProperties33 = new TableCellProperties();
					var tableCellWidth33 = new TableCellWidth { Width = "3417", Type = TableWidthUnitValues.Dxa };
					var shading1 = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

					tableCellProperties33.Append(tableCellWidth33);
					tableCellProperties33.Append(shading1);

					var paragraph37 = new Paragraph { RsidParagraphMarkRevision = "004950F8", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "12AB6F73", TextId = "77777777" };

					var paragraphProperties36 = new ParagraphProperties();
					var autoSpaceDE3 = new AutoSpaceDE { Val = false };
					var autoSpaceDN3 = new AutoSpaceDN { Val = false };
					var adjustRightIndent3 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines34 = new SpacingBetweenLines { After = "0" };
					var justification34 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
					var runFonts26 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize26 = new FontSize { Val = "20" };
					var fontSizeComplexScript4 = new FontSizeComplexScript { Val = "20" };
					var underline1 = new Underline { Val = UnderlineValues.Single };
					var languages4 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties3.Append(runFonts26);
					paragraphMarkRunProperties3.Append(fontSize26);
					paragraphMarkRunProperties3.Append(fontSizeComplexScript4);
					paragraphMarkRunProperties3.Append(underline1);
					paragraphMarkRunProperties3.Append(languages4);

					paragraphProperties36.Append(autoSpaceDE3);
					paragraphProperties36.Append(autoSpaceDN3);
					paragraphProperties36.Append(adjustRightIndent3);
					paragraphProperties36.Append(spacingBetweenLines34);
					paragraphProperties36.Append(justification34);
					paragraphProperties36.Append(paragraphMarkRunProperties3);

					paragraph37.Append(paragraphProperties36);

					var paragraph38 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "55DF56A2", TextId = "77777777" };

					var paragraphProperties37 = new ParagraphProperties();
					var autoSpaceDE4 = new AutoSpaceDE { Val = false };
					var autoSpaceDN4 = new AutoSpaceDN { Val = false };
					var adjustRightIndent4 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines35 = new SpacingBetweenLines { After = "0" };
					var justification35 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
					var runFonts27 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize27 = new FontSize { Val = "20" };
					var fontSizeComplexScript5 = new FontSizeComplexScript { Val = "20" };
					var underline2 = new Underline { Val = UnderlineValues.Single };
					var languages5 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties4.Append(runFonts27);
					paragraphMarkRunProperties4.Append(fontSize27);
					paragraphMarkRunProperties4.Append(fontSizeComplexScript5);
					paragraphMarkRunProperties4.Append(underline2);
					paragraphMarkRunProperties4.Append(languages5);

					paragraphProperties37.Append(autoSpaceDE4);
					paragraphProperties37.Append(autoSpaceDN4);
					paragraphProperties37.Append(adjustRightIndent4);
					paragraphProperties37.Append(spacingBetweenLines35);
					paragraphProperties37.Append(justification35);
					paragraphProperties37.Append(paragraphMarkRunProperties4);

					var run23 = new Run { RsidRunProperties = "00232403" };

					var runProperties23 = new RunProperties();
					var runFonts28 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize28 = new FontSize { Val = "20" };
					var fontSizeComplexScript6 = new FontSizeComplexScript { Val = "20" };
					var languages6 = new Languages { EastAsia = "en-US" };

					runProperties23.Append(runFonts28);
					runProperties23.Append(fontSize28);
					runProperties23.Append(fontSizeComplexScript6);
					runProperties23.Append(languages6);
					var text22 = new Text { Text = "______" };

					run23.Append(runProperties23);
					run23.Append(text22);

					var run24 = new Run { RsidRunProperties = "00232403" };

					var runProperties24 = new RunProperties();
					var runFonts29 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize29 = new FontSize { Val = "20" };
					var fontSizeComplexScript7 = new FontSizeComplexScript { Val = "20" };
					var underline3 = new Underline { Val = UnderlineValues.Single };
					var languages7 = new Languages { EastAsia = "en-US" };

					runProperties24.Append(runFonts29);
					runProperties24.Append(fontSize29);
					runProperties24.Append(fontSizeComplexScript7);
					runProperties24.Append(underline3);
					runProperties24.Append(languages7);
					var text23 = new Text { Text = "Главный бухгалтер" };

					run24.Append(runProperties24);
					run24.Append(text23);

					var run25 = new Run { RsidRunProperties = "00232403" };

					var runProperties25 = new RunProperties();
					var runFonts30 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize30 = new FontSize { Val = "20" };
					var fontSizeComplexScript8 = new FontSizeComplexScript { Val = "20" };
					var languages8 = new Languages { EastAsia = "en-US" };

					runProperties25.Append(runFonts30);
					runProperties25.Append(fontSize30);
					runProperties25.Append(fontSizeComplexScript8);
					runProperties25.Append(languages8);
					var text24 = new Text { Text = "_______" };

					run25.Append(runProperties25);
					run25.Append(text24);

					paragraph38.Append(paragraphProperties37);
					paragraph38.Append(run23);
					paragraph38.Append(run24);
					paragraph38.Append(run25);

					var paragraph39 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "740B661B", TextId = "77777777" };

					var paragraphProperties38 = new ParagraphProperties();
					var autoSpaceDE5 = new AutoSpaceDE { Val = false };
					var autoSpaceDN5 = new AutoSpaceDN { Val = false };
					var adjustRightIndent5 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines36 = new SpacingBetweenLines { After = "0" };
					var justification36 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
					var runFonts31 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize31 = new FontSize { Val = "20" };
					var fontSizeComplexScript9 = new FontSizeComplexScript { Val = "20" };
					var underline4 = new Underline { Val = UnderlineValues.Single };
					var languages9 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties5.Append(runFonts31);
					paragraphMarkRunProperties5.Append(fontSize31);
					paragraphMarkRunProperties5.Append(fontSizeComplexScript9);
					paragraphMarkRunProperties5.Append(underline4);
					paragraphMarkRunProperties5.Append(languages9);

					paragraphProperties38.Append(autoSpaceDE5);
					paragraphProperties38.Append(autoSpaceDN5);
					paragraphProperties38.Append(adjustRightIndent5);
					paragraphProperties38.Append(spacingBetweenLines36);
					paragraphProperties38.Append(justification36);
					paragraphProperties38.Append(paragraphMarkRunProperties5);

					paragraph39.Append(paragraphProperties38);

					tableCell33.Append(tableCellProperties33);
					tableCell33.Append(paragraph37);
					tableCell33.Append(paragraph38);
					tableCell33.Append(paragraph39);

					var tableCell34 = new TableCell();

					var tableCellProperties34 = new TableCellProperties();
					var tableCellWidth34 = new TableCellWidth { Width = "3417", Type = TableWidthUnitValues.Dxa };
					var shading2 = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

					tableCellProperties34.Append(tableCellWidth34);
					tableCellProperties34.Append(shading2);

					var paragraph40 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "7459112B", TextId = "77777777" };

					var paragraphProperties39 = new ParagraphProperties();
					var autoSpaceDE6 = new AutoSpaceDE { Val = false };
					var autoSpaceDN6 = new AutoSpaceDN { Val = false };
					var adjustRightIndent6 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines37 = new SpacingBetweenLines { After = "0" };
					var justification37 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
					var runFonts32 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize32 = new FontSize { Val = "20" };
					var fontSizeComplexScript10 = new FontSizeComplexScript { Val = "20" };
					var languages10 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties6.Append(runFonts32);
					paragraphMarkRunProperties6.Append(fontSize32);
					paragraphMarkRunProperties6.Append(fontSizeComplexScript10);
					paragraphMarkRunProperties6.Append(languages10);

					paragraphProperties39.Append(autoSpaceDE6);
					paragraphProperties39.Append(autoSpaceDN6);
					paragraphProperties39.Append(adjustRightIndent6);
					paragraphProperties39.Append(spacingBetweenLines37);
					paragraphProperties39.Append(justification37);
					paragraphProperties39.Append(paragraphMarkRunProperties6);

					paragraph40.Append(paragraphProperties39);

					var paragraph41 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "17F4861A", TextId = "77777777" };

					var paragraphProperties40 = new ParagraphProperties();
					var autoSpaceDE7 = new AutoSpaceDE { Val = false };
					var autoSpaceDN7 = new AutoSpaceDN { Val = false };
					var adjustRightIndent7 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines38 = new SpacingBetweenLines { After = "0" };
					var justification38 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
					var runFonts33 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize33 = new FontSize { Val = "20" };
					var fontSizeComplexScript11 = new FontSizeComplexScript { Val = "20" };
					var languages11 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties7.Append(runFonts33);
					paragraphMarkRunProperties7.Append(fontSize33);
					paragraphMarkRunProperties7.Append(fontSizeComplexScript11);
					paragraphMarkRunProperties7.Append(languages11);

					paragraphProperties40.Append(autoSpaceDE7);
					paragraphProperties40.Append(autoSpaceDN7);
					paragraphProperties40.Append(adjustRightIndent7);
					paragraphProperties40.Append(spacingBetweenLines38);
					paragraphProperties40.Append(justification38);
					paragraphProperties40.Append(paragraphMarkRunProperties7);

					var run26 = new Run { RsidRunProperties = "00232403" };

					var runProperties26 = new RunProperties();
					var runFonts34 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize34 = new FontSize { Val = "20" };
					var fontSizeComplexScript12 = new FontSizeComplexScript { Val = "20" };
					var languages12 = new Languages { EastAsia = "en-US" };

					runProperties26.Append(runFonts34);
					runProperties26.Append(fontSize34);
					runProperties26.Append(fontSizeComplexScript12);
					runProperties26.Append(languages12);
					var text25 = new Text { Space = SpaceProcessingModeValues.Preserve, Text = "________________________ " };

					run26.Append(runProperties26);
					run26.Append(text25);

					paragraph41.Append(paragraphProperties40);
					paragraph41.Append(run26);

					var paragraph42 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "1AC94114", TextId = "77777777" };

					var paragraphProperties41 = new ParagraphProperties();
					var autoSpaceDE8 = new AutoSpaceDE { Val = false };
					var autoSpaceDN8 = new AutoSpaceDN { Val = false };
					var adjustRightIndent8 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines39 = new SpacingBetweenLines { After = "0" };
					var justification39 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
					var runFonts35 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize35 = new FontSize { Val = "20" };
					var fontSizeComplexScript13 = new FontSizeComplexScript { Val = "20" };
					var languages13 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties8.Append(runFonts35);
					paragraphMarkRunProperties8.Append(fontSize35);
					paragraphMarkRunProperties8.Append(fontSizeComplexScript13);
					paragraphMarkRunProperties8.Append(languages13);

					paragraphProperties41.Append(autoSpaceDE8);
					paragraphProperties41.Append(autoSpaceDN8);
					paragraphProperties41.Append(adjustRightIndent8);
					paragraphProperties41.Append(spacingBetweenLines39);
					paragraphProperties41.Append(justification39);
					paragraphProperties41.Append(paragraphMarkRunProperties8);

					var run27 = new Run { RsidRunProperties = "00232403" };

					var runProperties27 = new RunProperties();
					var runFonts36 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize36 = new FontSize { Val = "20" };
					var fontSizeComplexScript14 = new FontSizeComplexScript { Val = "20" };
					var languages14 = new Languages { EastAsia = "en-US" };

					runProperties27.Append(runFonts36);
					runProperties27.Append(fontSize36);
					runProperties27.Append(fontSizeComplexScript14);
					runProperties27.Append(languages14);
					var text26 = new Text { Text = "(подпись)" };

					run27.Append(runProperties27);
					run27.Append(text26);

					paragraph42.Append(paragraphProperties41);
					paragraph42.Append(run27);

					tableCell34.Append(tableCellProperties34);
					tableCell34.Append(paragraph40);
					tableCell34.Append(paragraph41);
					tableCell34.Append(paragraph42);

					var tableCell35 = new TableCell();

					var tableCellProperties35 = new TableCellProperties();
					var tableCellWidth35 = new TableCellWidth { Width = "2916", Type = TableWidthUnitValues.Dxa };
					var shading3 = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

					tableCellProperties35.Append(tableCellWidth35);
					tableCellProperties35.Append(shading3);

					var paragraph43 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "35D2677F", TextId = "77777777" };

					var paragraphProperties42 = new ParagraphProperties();
					var autoSpaceDE9 = new AutoSpaceDE { Val = false };
					var autoSpaceDN9 = new AutoSpaceDN { Val = false };
					var adjustRightIndent9 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines40 = new SpacingBetweenLines { After = "0" };
					var justification40 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
					var runFonts37 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize37 = new FontSize { Val = "20" };
					var fontSizeComplexScript15 = new FontSizeComplexScript { Val = "20" };
					var languages15 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties9.Append(runFonts37);
					paragraphMarkRunProperties9.Append(fontSize37);
					paragraphMarkRunProperties9.Append(fontSizeComplexScript15);
					paragraphMarkRunProperties9.Append(languages15);

					paragraphProperties42.Append(autoSpaceDE9);
					paragraphProperties42.Append(autoSpaceDN9);
					paragraphProperties42.Append(adjustRightIndent9);
					paragraphProperties42.Append(spacingBetweenLines40);
					paragraphProperties42.Append(justification40);
					paragraphProperties42.Append(paragraphMarkRunProperties9);

					paragraph43.Append(paragraphProperties42);

					var paragraph44 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "63D934B2", TextId = "77777777" };

					var paragraphProperties43 = new ParagraphProperties();
					var autoSpaceDE10 = new AutoSpaceDE { Val = false };
					var autoSpaceDN10 = new AutoSpaceDN { Val = false };
					var adjustRightIndent10 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines41 = new SpacingBetweenLines { After = "0" };
					var justification41 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
					var runFonts38 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize38 = new FontSize { Val = "20" };
					var fontSizeComplexScript16 = new FontSizeComplexScript { Val = "20" };
					var languages16 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties10.Append(runFonts38);
					paragraphMarkRunProperties10.Append(fontSize38);
					paragraphMarkRunProperties10.Append(fontSizeComplexScript16);
					paragraphMarkRunProperties10.Append(languages16);

					paragraphProperties43.Append(autoSpaceDE10);
					paragraphProperties43.Append(autoSpaceDN10);
					paragraphProperties43.Append(adjustRightIndent10);
					paragraphProperties43.Append(spacingBetweenLines41);
					paragraphProperties43.Append(justification41);
					paragraphProperties43.Append(paragraphMarkRunProperties10);

					var run28 = new Run { RsidRunProperties = "00232403" };

					var runProperties28 = new RunProperties();
					var runFonts39 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize39 = new FontSize { Val = "20" };
					var fontSizeComplexScript17 = new FontSizeComplexScript { Val = "20" };
					var languages17 = new Languages { EastAsia = "en-US" };

					runProperties28.Append(runFonts39);
					runProperties28.Append(fontSize39);
					runProperties28.Append(fontSizeComplexScript17);
					runProperties28.Append(languages17);
					var text27 = new Text { Text = "___________________________" };

					run28.Append(runProperties28);
					run28.Append(text27);

					paragraph44.Append(paragraphProperties43);
					paragraph44.Append(run28);

					var paragraph45 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "1EC78E65", TextId = "77777777" };

					var paragraphProperties44 = new ParagraphProperties();
					var autoSpaceDE11 = new AutoSpaceDE { Val = false };
					var autoSpaceDN11 = new AutoSpaceDN { Val = false };
					var adjustRightIndent11 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines42 = new SpacingBetweenLines { After = "0" };
					var justification42 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
					var runFonts40 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize40 = new FontSize { Val = "20" };
					var fontSizeComplexScript18 = new FontSizeComplexScript { Val = "20" };
					var languages18 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties11.Append(runFonts40);
					paragraphMarkRunProperties11.Append(fontSize40);
					paragraphMarkRunProperties11.Append(fontSizeComplexScript18);
					paragraphMarkRunProperties11.Append(languages18);

					paragraphProperties44.Append(autoSpaceDE11);
					paragraphProperties44.Append(autoSpaceDN11);
					paragraphProperties44.Append(adjustRightIndent11);
					paragraphProperties44.Append(spacingBetweenLines42);
					paragraphProperties44.Append(justification42);
					paragraphProperties44.Append(paragraphMarkRunProperties11);

					var run29 = new Run { RsidRunProperties = "00232403" };

					var runProperties29 = new RunProperties();
					var runFonts41 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize41 = new FontSize { Val = "20" };
					var fontSizeComplexScript19 = new FontSizeComplexScript { Val = "20" };
					var languages19 = new Languages { EastAsia = "en-US" };

					runProperties29.Append(runFonts41);
					runProperties29.Append(fontSize41);
					runProperties29.Append(fontSizeComplexScript19);
					runProperties29.Append(languages19);
					var text28 = new Text { Text = "(фамилия, инициалы)" };

					run29.Append(runProperties29);
					run29.Append(text28);

					paragraph45.Append(paragraphProperties44);
					paragraph45.Append(run29);

					tableCell35.Append(tableCellProperties35);
					tableCell35.Append(paragraph43);
					tableCell35.Append(paragraph44);
					tableCell35.Append(paragraph45);

					tableRow5.Append(tableCell33);
					tableRow5.Append(tableCell34);
					tableRow5.Append(tableCell35);

					TableRow tableRow6 = new TableRow { RsidTableRowMarkRevision = "00232403", RsidTableRowAddition = "004950F8", RsidTableRowProperties = "00751945", ParagraphId = "5E09C7EA", TextId = "77777777" };

					var tableCell36 = new TableCell();

					var tableCellProperties36 = new TableCellProperties();
					var tableCellWidth36 = new TableCellWidth { Width = "3417", Type = TableWidthUnitValues.Dxa };
					var shading4 = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

					tableCellProperties36.Append(tableCellWidth36);
					tableCellProperties36.Append(shading4);

					var paragraph46 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "0B078ACA", TextId = "77777777" };

					var paragraphProperties45 = new ParagraphProperties();
					var autoSpaceDE12 = new AutoSpaceDE { Val = false };
					var autoSpaceDN12 = new AutoSpaceDN { Val = false };
					var adjustRightIndent12 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines43 = new SpacingBetweenLines { After = "0" };

					var paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
					var runFonts42 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize42 = new FontSize { Val = "20" };
					var fontSizeComplexScript20 = new FontSizeComplexScript { Val = "20" };
					var languages20 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties12.Append(runFonts42);
					paragraphMarkRunProperties12.Append(fontSize42);
					paragraphMarkRunProperties12.Append(fontSizeComplexScript20);
					paragraphMarkRunProperties12.Append(languages20);

					paragraphProperties45.Append(autoSpaceDE12);
					paragraphProperties45.Append(autoSpaceDN12);
					paragraphProperties45.Append(adjustRightIndent12);
					paragraphProperties45.Append(spacingBetweenLines43);
					paragraphProperties45.Append(paragraphMarkRunProperties12);

					paragraph46.Append(paragraphProperties45);

					var paragraph47 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "3B64293F", TextId = "77777777" };

					var paragraphProperties46 = new ParagraphProperties();
					var autoSpaceDE13 = new AutoSpaceDE { Val = false };
					var autoSpaceDN13 = new AutoSpaceDN { Val = false };
					var adjustRightIndent13 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines44 = new SpacingBetweenLines { After = "0" };

					var paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
					var runFonts43 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize43 = new FontSize { Val = "20" };
					var fontSizeComplexScript21 = new FontSizeComplexScript { Val = "20" };
					var languages21 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties13.Append(runFonts43);
					paragraphMarkRunProperties13.Append(fontSize43);
					paragraphMarkRunProperties13.Append(fontSizeComplexScript21);
					paragraphMarkRunProperties13.Append(languages21);

					paragraphProperties46.Append(autoSpaceDE13);
					paragraphProperties46.Append(autoSpaceDN13);
					paragraphProperties46.Append(adjustRightIndent13);
					paragraphProperties46.Append(spacingBetweenLines44);
					paragraphProperties46.Append(paragraphMarkRunProperties13);

					paragraph47.Append(paragraphProperties46);

					var paragraph48 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "0C8F22C4", TextId = "77777777" };

					var paragraphProperties47 = new ParagraphProperties();
					var autoSpaceDE14 = new AutoSpaceDE { Val = false };
					var autoSpaceDN14 = new AutoSpaceDN { Val = false };
					var adjustRightIndent14 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines45 = new SpacingBetweenLines { After = "0" };
					var justification43 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
					var runFonts44 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize44 = new FontSize { Val = "20" };
					var fontSizeComplexScript22 = new FontSizeComplexScript { Val = "20" };
					var languages22 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties14.Append(runFonts44);
					paragraphMarkRunProperties14.Append(fontSize44);
					paragraphMarkRunProperties14.Append(fontSizeComplexScript22);
					paragraphMarkRunProperties14.Append(languages22);

					paragraphProperties47.Append(autoSpaceDE14);
					paragraphProperties47.Append(autoSpaceDN14);
					paragraphProperties47.Append(adjustRightIndent14);
					paragraphProperties47.Append(spacingBetweenLines45);
					paragraphProperties47.Append(justification43);
					paragraphProperties47.Append(paragraphMarkRunProperties14);

					var run30 = new Run { RsidRunProperties = "00232403" };

					var runProperties30 = new RunProperties();
					var runFonts45 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize45 = new FontSize { Val = "20" };
					var fontSizeComplexScript23 = new FontSizeComplexScript { Val = "20" };
					var languages23 = new Languages { EastAsia = "en-US" };

					runProperties30.Append(runFonts45);
					runProperties30.Append(fontSize45);
					runProperties30.Append(fontSizeComplexScript23);
					runProperties30.Append(languages23);
					var text29 = new Text { Text = "____________________________ (наименование должности руководителя Организации)" };

					run30.Append(runProperties30);
					run30.Append(text29);

					paragraph48.Append(paragraphProperties47);
					paragraph48.Append(run30);

					var paragraph49 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "004950F8", RsidRunAdditionDefault = "004950F8", ParagraphId = "6A357A9D", TextId = "77777777" };

					var paragraphProperties48 = new ParagraphProperties();
					var autoSpaceDE15 = new AutoSpaceDE { Val = false };
					var autoSpaceDN15 = new AutoSpaceDN { Val = false };
					var adjustRightIndent15 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines46 = new SpacingBetweenLines { After = "0" };

					var paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
					var runFonts46 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize46 = new FontSize { Val = "20" };
					var fontSizeComplexScript24 = new FontSizeComplexScript { Val = "20" };
					var languages24 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties15.Append(runFonts46);
					paragraphMarkRunProperties15.Append(fontSize46);
					paragraphMarkRunProperties15.Append(fontSizeComplexScript24);
					paragraphMarkRunProperties15.Append(languages24);

					paragraphProperties48.Append(autoSpaceDE15);
					paragraphProperties48.Append(autoSpaceDN15);
					paragraphProperties48.Append(adjustRightIndent15);
					paragraphProperties48.Append(spacingBetweenLines46);
					paragraphProperties48.Append(paragraphMarkRunProperties15);

					var run31 = new Run { RsidRunProperties = "00232403" };

					var runProperties31 = new RunProperties();
					var runFonts47 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize47 = new FontSize { Val = "20" };
					var fontSizeComplexScript25 = new FontSizeComplexScript { Val = "20" };
					var languages25 = new Languages { EastAsia = "en-US" };

					runProperties31.Append(runFonts47);
					runProperties31.Append(fontSize47);
					runProperties31.Append(fontSizeComplexScript25);
					runProperties31.Append(languages25);
					var text30 = new Text { Text = "М.П." };

					run31.Append(runProperties31);
					run31.Append(text30);

					paragraph49.Append(paragraphProperties48);
					paragraph49.Append(run31);

					var paragraph50 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "20ADC8D5", TextId = "77777777" };

					var paragraphProperties49 = new ParagraphProperties();
					var autoSpaceDE16 = new AutoSpaceDE { Val = false };
					var autoSpaceDN16 = new AutoSpaceDN { Val = false };
					var adjustRightIndent16 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines47 = new SpacingBetweenLines { After = "0" };
					var justification44 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
					var runFonts48 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize48 = new FontSize { Val = "20" };
					var fontSizeComplexScript26 = new FontSizeComplexScript { Val = "20" };
					var languages26 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties16.Append(runFonts48);
					paragraphMarkRunProperties16.Append(fontSize48);
					paragraphMarkRunProperties16.Append(fontSizeComplexScript26);
					paragraphMarkRunProperties16.Append(languages26);

					paragraphProperties49.Append(autoSpaceDE16);
					paragraphProperties49.Append(autoSpaceDN16);
					paragraphProperties49.Append(adjustRightIndent16);
					paragraphProperties49.Append(spacingBetweenLines47);
					paragraphProperties49.Append(justification44);
					paragraphProperties49.Append(paragraphMarkRunProperties16);

					paragraph50.Append(paragraphProperties49);

					tableCell36.Append(tableCellProperties36);
					tableCell36.Append(paragraph46);
					tableCell36.Append(paragraph47);
					tableCell36.Append(paragraph48);
					tableCell36.Append(paragraph49);
					tableCell36.Append(paragraph50);

					var tableCell37 = new TableCell();

					var tableCellProperties37 = new TableCellProperties();
					var tableCellWidth37 = new TableCellWidth { Width = "3417", Type = TableWidthUnitValues.Dxa };
					var shading5 = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

					tableCellProperties37.Append(tableCellWidth37);
					tableCellProperties37.Append(shading5);

					var paragraph51 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "45FE4728", TextId = "77777777" };

					var paragraphProperties50 = new ParagraphProperties();
					var autoSpaceDE17 = new AutoSpaceDE { Val = false };
					var autoSpaceDN17 = new AutoSpaceDN { Val = false };
					var adjustRightIndent17 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines48 = new SpacingBetweenLines { After = "0" };
					var justification45 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
					var runFonts49 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize49 = new FontSize { Val = "20" };
					var fontSizeComplexScript27 = new FontSizeComplexScript { Val = "20" };
					var languages27 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties17.Append(runFonts49);
					paragraphMarkRunProperties17.Append(fontSize49);
					paragraphMarkRunProperties17.Append(fontSizeComplexScript27);
					paragraphMarkRunProperties17.Append(languages27);

					paragraphProperties50.Append(autoSpaceDE17);
					paragraphProperties50.Append(autoSpaceDN17);
					paragraphProperties50.Append(adjustRightIndent17);
					paragraphProperties50.Append(spacingBetweenLines48);
					paragraphProperties50.Append(justification45);
					paragraphProperties50.Append(paragraphMarkRunProperties17);

					paragraph51.Append(paragraphProperties50);

					var paragraph52 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "46FC90B7", TextId = "77777777" };

					var paragraphProperties51 = new ParagraphProperties();
					var autoSpaceDE18 = new AutoSpaceDE { Val = false };
					var autoSpaceDN18 = new AutoSpaceDN { Val = false };
					var adjustRightIndent18 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines49 = new SpacingBetweenLines { After = "0" };
					var justification46 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
					var runFonts50 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize50 = new FontSize { Val = "20" };
					var fontSizeComplexScript28 = new FontSizeComplexScript { Val = "20" };
					var languages28 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties18.Append(runFonts50);
					paragraphMarkRunProperties18.Append(fontSize50);
					paragraphMarkRunProperties18.Append(fontSizeComplexScript28);
					paragraphMarkRunProperties18.Append(languages28);

					paragraphProperties51.Append(autoSpaceDE18);
					paragraphProperties51.Append(autoSpaceDN18);
					paragraphProperties51.Append(adjustRightIndent18);
					paragraphProperties51.Append(spacingBetweenLines49);
					paragraphProperties51.Append(justification46);
					paragraphProperties51.Append(paragraphMarkRunProperties18);

					paragraph52.Append(paragraphProperties51);

					var paragraph53 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "62488562", TextId = "77777777" };

					var paragraphProperties52 = new ParagraphProperties();
					var autoSpaceDE19 = new AutoSpaceDE { Val = false };
					var autoSpaceDN19 = new AutoSpaceDN { Val = false };
					var adjustRightIndent19 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines50 = new SpacingBetweenLines { After = "0" };
					var justification47 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
					var runFonts51 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize51 = new FontSize { Val = "20" };
					var fontSizeComplexScript29 = new FontSizeComplexScript { Val = "20" };
					var languages29 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties19.Append(runFonts51);
					paragraphMarkRunProperties19.Append(fontSize51);
					paragraphMarkRunProperties19.Append(fontSizeComplexScript29);
					paragraphMarkRunProperties19.Append(languages29);

					paragraphProperties52.Append(autoSpaceDE19);
					paragraphProperties52.Append(autoSpaceDN19);
					paragraphProperties52.Append(adjustRightIndent19);
					paragraphProperties52.Append(spacingBetweenLines50);
					paragraphProperties52.Append(justification47);
					paragraphProperties52.Append(paragraphMarkRunProperties19);

					var run32 = new Run { RsidRunProperties = "00232403" };

					var runProperties32 = new RunProperties();
					var runFonts52 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize52 = new FontSize { Val = "20" };
					var fontSizeComplexScript30 = new FontSizeComplexScript { Val = "20" };
					var languages30 = new Languages { EastAsia = "en-US" };

					runProperties32.Append(runFonts52);
					runProperties32.Append(fontSize52);
					runProperties32.Append(fontSizeComplexScript30);
					runProperties32.Append(languages30);
					var text31 = new Text { Space = SpaceProcessingModeValues.Preserve, Text = "________________________ " };

					run32.Append(runProperties32);
					run32.Append(text31);

					paragraph53.Append(paragraphProperties52);
					paragraph53.Append(run32);

					var paragraph54 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "1331F6B1", TextId = "77777777" };

					var paragraphProperties53 = new ParagraphProperties();
					var autoSpaceDE20 = new AutoSpaceDE { Val = false };
					var autoSpaceDN20 = new AutoSpaceDN { Val = false };
					var adjustRightIndent20 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines51 = new SpacingBetweenLines { After = "0" };
					var justification48 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
					var runFonts53 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize53 = new FontSize { Val = "20" };
					var fontSizeComplexScript31 = new FontSizeComplexScript { Val = "20" };
					var languages31 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties20.Append(runFonts53);
					paragraphMarkRunProperties20.Append(fontSize53);
					paragraphMarkRunProperties20.Append(fontSizeComplexScript31);
					paragraphMarkRunProperties20.Append(languages31);

					paragraphProperties53.Append(autoSpaceDE20);
					paragraphProperties53.Append(autoSpaceDN20);
					paragraphProperties53.Append(adjustRightIndent20);
					paragraphProperties53.Append(spacingBetweenLines51);
					paragraphProperties53.Append(justification48);
					paragraphProperties53.Append(paragraphMarkRunProperties20);

					var run33 = new Run { RsidRunProperties = "00232403" };

					var runProperties33 = new RunProperties();
					var runFonts54 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize54 = new FontSize { Val = "20" };
					var fontSizeComplexScript32 = new FontSizeComplexScript { Val = "20" };
					var languages32 = new Languages { EastAsia = "en-US" };

					runProperties33.Append(runFonts54);
					runProperties33.Append(fontSize54);
					runProperties33.Append(fontSizeComplexScript32);
					runProperties33.Append(languages32);
					var text32 = new Text { Text = "(подпись)" };

					run33.Append(runProperties33);
					run33.Append(text32);

					paragraph54.Append(paragraphProperties53);
					paragraph54.Append(run33);

					tableCell37.Append(tableCellProperties37);
					tableCell37.Append(paragraph51);
					tableCell37.Append(paragraph52);
					tableCell37.Append(paragraph53);
					tableCell37.Append(paragraph54);

					var tableCell38 = new TableCell();

					var tableCellProperties38 = new TableCellProperties();
					var tableCellWidth38 = new TableCellWidth { Width = "2916", Type = TableWidthUnitValues.Dxa };
					var shading6 = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

					tableCellProperties38.Append(tableCellWidth38);
					tableCellProperties38.Append(shading6);

					var paragraph55 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "7ED1A1F7", TextId = "77777777" };

					var paragraphProperties54 = new ParagraphProperties();
					var autoSpaceDE21 = new AutoSpaceDE { Val = false };
					var autoSpaceDN21 = new AutoSpaceDN { Val = false };
					var adjustRightIndent21 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines52 = new SpacingBetweenLines { After = "0" };

					var paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
					var runFonts55 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize55 = new FontSize { Val = "20" };
					var fontSizeComplexScript33 = new FontSizeComplexScript { Val = "20" };
					var languages33 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties21.Append(runFonts55);
					paragraphMarkRunProperties21.Append(fontSize55);
					paragraphMarkRunProperties21.Append(fontSizeComplexScript33);
					paragraphMarkRunProperties21.Append(languages33);

					paragraphProperties54.Append(autoSpaceDE21);
					paragraphProperties54.Append(autoSpaceDN21);
					paragraphProperties54.Append(adjustRightIndent21);
					paragraphProperties54.Append(spacingBetweenLines52);
					paragraphProperties54.Append(paragraphMarkRunProperties21);

					paragraph55.Append(paragraphProperties54);

					var paragraph56 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "4B3F6D60", TextId = "77777777" };

					var paragraphProperties55 = new ParagraphProperties();
					var autoSpaceDE22 = new AutoSpaceDE { Val = false };
					var autoSpaceDN22 = new AutoSpaceDN { Val = false };
					var adjustRightIndent22 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines53 = new SpacingBetweenLines { After = "0" };

					var paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
					var runFonts56 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize56 = new FontSize { Val = "20" };
					var fontSizeComplexScript34 = new FontSizeComplexScript { Val = "20" };
					var languages34 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties22.Append(runFonts56);
					paragraphMarkRunProperties22.Append(fontSize56);
					paragraphMarkRunProperties22.Append(fontSizeComplexScript34);
					paragraphMarkRunProperties22.Append(languages34);

					paragraphProperties55.Append(autoSpaceDE22);
					paragraphProperties55.Append(autoSpaceDN22);
					paragraphProperties55.Append(adjustRightIndent22);
					paragraphProperties55.Append(spacingBetweenLines53);
					paragraphProperties55.Append(paragraphMarkRunProperties22);

					paragraph56.Append(paragraphProperties55);

					var paragraph57 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "5B53E466", TextId = "77777777" };

					var paragraphProperties56 = new ParagraphProperties();
					var autoSpaceDE23 = new AutoSpaceDE { Val = false };
					var autoSpaceDN23 = new AutoSpaceDN { Val = false };
					var adjustRightIndent23 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines54 = new SpacingBetweenLines { After = "0" };
					var justification49 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
					var runFonts57 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize57 = new FontSize { Val = "20" };
					var fontSizeComplexScript35 = new FontSizeComplexScript { Val = "20" };
					var languages35 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties23.Append(runFonts57);
					paragraphMarkRunProperties23.Append(fontSize57);
					paragraphMarkRunProperties23.Append(fontSizeComplexScript35);
					paragraphMarkRunProperties23.Append(languages35);

					paragraphProperties56.Append(autoSpaceDE23);
					paragraphProperties56.Append(autoSpaceDN23);
					paragraphProperties56.Append(adjustRightIndent23);
					paragraphProperties56.Append(spacingBetweenLines54);
					paragraphProperties56.Append(justification49);
					paragraphProperties56.Append(paragraphMarkRunProperties23);

					var run34 = new Run { RsidRunProperties = "00232403" };

					var runProperties34 = new RunProperties();
					var runFonts58 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize58 = new FontSize { Val = "20" };
					var fontSizeComplexScript36 = new FontSizeComplexScript { Val = "20" };
					var languages36 = new Languages { EastAsia = "en-US" };

					runProperties34.Append(runFonts58);
					runProperties34.Append(fontSize58);
					runProperties34.Append(fontSizeComplexScript36);
					runProperties34.Append(languages36);
					var text33 = new Text { Text = "___________________________" };

					run34.Append(runProperties34);
					run34.Append(text33);

					paragraph57.Append(paragraphProperties56);
					paragraph57.Append(run34);

					var paragraph58 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "3C6E05CE", TextId = "77777777" };

					var paragraphProperties57 = new ParagraphProperties();
					var autoSpaceDE24 = new AutoSpaceDE { Val = false };
					var autoSpaceDN24 = new AutoSpaceDN { Val = false };
					var adjustRightIndent24 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines55 = new SpacingBetweenLines { After = "0" };
					var justification50 = new Justification { Val = JustificationValues.Center };

					var paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
					var runFonts59 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize59 = new FontSize { Val = "20" };
					var fontSizeComplexScript37 = new FontSizeComplexScript { Val = "20" };
					var languages37 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties24.Append(runFonts59);
					paragraphMarkRunProperties24.Append(fontSize59);
					paragraphMarkRunProperties24.Append(fontSizeComplexScript37);
					paragraphMarkRunProperties24.Append(languages37);

					paragraphProperties57.Append(autoSpaceDE24);
					paragraphProperties57.Append(autoSpaceDN24);
					paragraphProperties57.Append(adjustRightIndent24);
					paragraphProperties57.Append(spacingBetweenLines55);
					paragraphProperties57.Append(justification50);
					paragraphProperties57.Append(paragraphMarkRunProperties24);

					var run35 = new Run { RsidRunProperties = "00232403" };

					var runProperties35 = new RunProperties();
					var runFonts60 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize60 = new FontSize { Val = "20" };
					var fontSizeComplexScript38 = new FontSizeComplexScript { Val = "20" };
					var languages38 = new Languages { EastAsia = "en-US" };

					runProperties35.Append(runFonts60);
					runProperties35.Append(fontSize60);
					runProperties35.Append(fontSizeComplexScript38);
					runProperties35.Append(languages38);
					var text34 = new Text { Text = "(фамилия, инициалы)" };

					run35.Append(runProperties35);
					run35.Append(text34);

					paragraph58.Append(paragraphProperties57);
					paragraph58.Append(run35);

					var paragraph59 = new Paragraph { RsidParagraphMarkRevision = "00232403", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "00751945", RsidRunAdditionDefault = "004950F8", ParagraphId = "63E18720", TextId = "77777777" };

					var paragraphProperties58 = new ParagraphProperties();
					var autoSpaceDE25 = new AutoSpaceDE { Val = false };
					var autoSpaceDN25 = new AutoSpaceDN { Val = false };
					var adjustRightIndent25 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines56 = new SpacingBetweenLines { After = "0" };

					var paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
					var runFonts61 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize61 = new FontSize { Val = "20" };
					var fontSizeComplexScript39 = new FontSizeComplexScript { Val = "20" };
					var languages39 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties25.Append(runFonts61);
					paragraphMarkRunProperties25.Append(fontSize61);
					paragraphMarkRunProperties25.Append(fontSizeComplexScript39);
					paragraphMarkRunProperties25.Append(languages39);

					paragraphProperties58.Append(autoSpaceDE25);
					paragraphProperties58.Append(autoSpaceDN25);
					paragraphProperties58.Append(adjustRightIndent25);
					paragraphProperties58.Append(spacingBetweenLines56);
					paragraphProperties58.Append(paragraphMarkRunProperties25);

					paragraph59.Append(paragraphProperties58);

					tableCell38.Append(tableCellProperties38);
					tableCell38.Append(paragraph55);
					tableCell38.Append(paragraph56);
					tableCell38.Append(paragraph57);
					tableCell38.Append(paragraph58);
					tableCell38.Append(paragraph59);

					tableRow6.Append(tableCell36);
					tableRow6.Append(tableCell37);
					tableRow6.Append(tableCell38);

					table2.Append(tableProperties2);
					table2.Append(tableGrid2);
					table2.Append(tableRow5);
					table2.Append(tableRow6);

					var paragraph60 = new Paragraph { RsidParagraphMarkRevision = "004950F8", RsidParagraphAddition = "004950F8", RsidParagraphProperties = "004950F8", RsidRunAdditionDefault = "004950F8", ParagraphId = "38474E6D", TextId = "77777777" };

					var paragraphProperties59 = new ParagraphProperties();
					var autoSpaceDE26 = new AutoSpaceDE { Val = false };
					var autoSpaceDN26 = new AutoSpaceDN { Val = false };
					var adjustRightIndent26 = new AdjustRightIndent { Val = false };
					var spacingBetweenLines57 = new SpacingBetweenLines { After = "0" };
					var justification51 = new Justification { Val = JustificationValues.Both };

					var paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
					var runFonts62 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize62 = new FontSize { Val = "20" };
					var fontSizeComplexScript40 = new FontSizeComplexScript { Val = "20" };
					var languages40 = new Languages { EastAsia = "en-US" };

					paragraphMarkRunProperties26.Append(runFonts62);
					paragraphMarkRunProperties26.Append(fontSize62);
					paragraphMarkRunProperties26.Append(fontSizeComplexScript40);
					paragraphMarkRunProperties26.Append(languages40);

					paragraphProperties59.Append(autoSpaceDE26);
					paragraphProperties59.Append(autoSpaceDN26);
					paragraphProperties59.Append(adjustRightIndent26);
					paragraphProperties59.Append(spacingBetweenLines57);
					paragraphProperties59.Append(justification51);
					paragraphProperties59.Append(paragraphMarkRunProperties26);

					var run36 = new Run { RsidRunProperties = "00232403" };

					var runProperties36 = new RunProperties();
					var runFonts63 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize63 = new FontSize { Val = "20" };
					var fontSizeComplexScript41 = new FontSizeComplexScript { Val = "20" };
					var languages41 = new Languages { EastAsia = "en-US" };

					runProperties36.Append(runFonts63);
					runProperties36.Append(fontSize63);
					runProperties36.Append(fontSizeComplexScript41);
					runProperties36.Append(languages41);
					var text35 = new Text { Text = "«___» __________ 20___ г." };

					run36.Append(runProperties36);
					run36.Append(text35);

					var run37 = new Run { RsidRunProperties = "004950F8" };

					var runProperties37 = new RunProperties();
					var runFonts64 = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
					var fontSize64 = new FontSize { Val = "20" };
					var fontSizeComplexScript42 = new FontSizeComplexScript { Val = "20" };
					var languages42 = new Languages { EastAsia = "en-US" };

					runProperties37.Append(runFonts64);
					runProperties37.Append(fontSize64);
					runProperties37.Append(fontSizeComplexScript42);
					runProperties37.Append(languages42);
					var text36 = new Text { Space = SpaceProcessingModeValues.Preserve, Text = _SPACE };

					run37.Append(runProperties37);
					run37.Append(text36);

					paragraph60.Append(paragraphProperties59);
					paragraph60.Append(run36);
					paragraph60.Append(run37);

					doc.AppendChild(table2);
					doc.AppendChild(paragraph60);
				}


				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Перечень" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		/// Блок подписией для ЛОК 2020
		/// </summary>
		public void SignBlockNotification2020(Document doc, Account account, string functionName = _SPACE)
		{
			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = _SIZE_22 });

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));
			var signRunProperties = new RunProperties().SetFont().SetFontSizeSupperscript();

			var table = new Table();
			table.AppendChild(tblProp.CloneNode(true));


			//отступ - нулевая строка
			{
				doc.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Both },
				new SpacingBetweenLines { After = _SIZE_20 })));
			}
			//первая строка
			{
				var row = new TableRow();

				var cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_1.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_2.ToString() },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_3.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "1" }), new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_4.ToString() },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Left }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_5.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "1" }), new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_6.ToString() },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(DateTime.Now.Date.FormatEx()))));
				row.AppendChild(cell);

				table.AppendChild(row);
			}

			//вторая строка
			{
				var row = new TableRow();

				var cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_1.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_2.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(signRunProperties.CloneNode(true), new Text("(подпись заявителя)"))));
				row.AppendChild(cell);


				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_3.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_4.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(signRunProperties.CloneNode(true), new Text("(ФИО заявителя)"))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_5.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_6.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(signRunProperties.CloneNode(true), new Text("(дата)"))));
				row.AppendChild(cell);

				table.AppendChild(row);
			}

			//третья строка
			{
				var row = new TableRow();

				var cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_1.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(functionName))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_2.ToString() },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(
							$"{account.Name.FormatEx()}{(string.IsNullOrWhiteSpace(account.Position) ? string.Empty : $", {account.Position}")}"))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_3.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "1" }), new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_4.ToString() },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Left }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_5.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "1" }), new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_6.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				table.AppendChild(row);
			}

			//четвёртая строка
			{
				var row = new TableRow();

				var cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_1.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_2.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(signRunProperties.CloneNode(true), new Text("(ФИО работника, должность)"))));
				row.AppendChild(cell);


				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_3.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_4.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(signRunProperties.CloneNode(true), new Text("(подпись работника)"))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_5.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_6.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(signRunProperties.CloneNode(true), new Text(_SPACE))));
				row.AppendChild(cell);

				table.AppendChild(row);
			}

			doc.AppendChild(table);
		}

		/// <summary>
		/// Блок подписией для ЛОК 2022
		/// </summary>
		public void SignBlockNotification2022(Document doc, Account account, string functionName = _SPACE)
		{
			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = _SIZE_22 });

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));
			var signRunProperties = new RunProperties().SetFont().SetFontSizeSupperscript();

			var table = new Table();
			table.AppendChild(tblProp.CloneNode(true));


			//отступ - нулевая строка
			{
				doc.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Both },
				new SpacingBetweenLines { After = _SIZE_20 })));
			}

			//третья строка
			{
				var row = new TableRow();

				var cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_1.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(functionName))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_2.ToString() },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(
							$"{account.Name.FormatEx()}{(string.IsNullOrWhiteSpace(account.Position) ? string.Empty : $", {account.Position}")}"))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_3.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "1" }), new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_4.ToString() },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Left }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_5.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "1" }), new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_6.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "1" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				table.AppendChild(row);
			}

			//четвёртая строка
			{
				var row = new TableRow();

				var cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_1.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_2.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(signRunProperties.CloneNode(true), new Text("(ФИО работника, должность)"))));
				row.AppendChild(cell);


				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_3.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_4.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(signRunProperties.CloneNode(true), new Text("(подпись работника)"))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_5.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_6.ToString() }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(signRunProperties.CloneNode(true), new Text(_SPACE))));
				row.AppendChild(cell);

				table.AppendChild(row);
			}

			doc.AppendChild(table);
		}

		/// <summary>
		/// Заявление (Request) в документе
		/// </summary>
		/// <returns></returns>
		public IDocument RequestDocument(Request request, ICollection<Booking> bookings)
		{
			if (request == null || request.IsDeleted) return null;

			var isNotYouthRest = request.TypeOfRestId != (long)TypeOfRestEnum.YouthRestOrphanCamps &&
								 request.TypeOfRestId != (long)TypeOfRestEnum.YouthRestCamps;

			MemoryStream ms;

			using (ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					mainPart.Document = doc;

					var section = new SectionProperties();
					section.AppendChild(new PageMargin { Bottom = 1000, Top = 1000, Left = 1000, Right = 1000 });
					mainPart.Document.Body.AppendChild(section);

					AppendHeader(doc);
					var generalBlock = new List<Tuple<string, string>>
					{
						new Tuple<string, string>("Номер заявления", request.RequestNumber.FormatEx()),
						new Tuple<string, string>("Дата заявления", request.DateRequest.FormatEx("dd.MM.yyyy")),
						new Tuple<string, string>("Источник заявления", request.Source.Name.FormatEx("-", false))
					};

					if (!string.IsNullOrEmpty(request.RequestNumberMpgu))
					{
						generalBlock.Add(new Tuple<string, string>("Номер заявления МПГУ", request.RequestNumberMpgu));
					}

					if (request.DeclineReason != null)
					{
						generalBlock.Add(new Tuple<string, string>("Причина отказа", request.DeclineReason.Name));
					}

					if (!string.IsNullOrWhiteSpace(request.StatusApplicant) && isNotYouthRest)
					{
						var statusApplicant = GetStatusApplicantName(request.StatusApplicant);

						if (!string.IsNullOrWhiteSpace(statusApplicant))
						{
							generalBlock.Add(new Tuple<string, string>("Заявление подает", statusApplicant));
						}
					}

					AppendBlock(doc, "Общие сведения о заявлении", generalBlock);

					var typeAndTypeRest = new List<Tuple<string, string>>
					{
						new Tuple<string, string>("Цель обращения",
							request.NullSafe(r => r.TypeOfRest.Name.FormatEx("-", false)))
					};

					if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation &&
						request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest &&
						request.TypeOfRestId != (long)TypeOfRestEnum.MoneyOn3To7 &&
						request.TypeOfRestId != (long)TypeOfRestEnum.MoneyOn7To15)
					{
						if (request.IsFirstCompany)
						{
							var firstTimeOfRest = request.TimesOfRest?.Where(ss => ss.Order > 0)
								.Select(ss => ss.TimeOfRest).FirstOrDefault();
							var secondTimeOfRest = request.TimesOfRest?.Where(ss => ss.Order > 0)
								.Select(ss => ss.TimeOfRest).LastOrDefault();
							if (secondTimeOfRest?.Id == firstTimeOfRest?.Id)
							{
								secondTimeOfRest = null;
							}

							typeAndTypeRest.Add(new Tuple<string, string>("Приоритетное время отдыха",
								request.TimeOfRest?.Name.FormatEx("-", false)));
							typeAndTypeRest.Add(new Tuple<string, string>("Дополнительные времена отдыха",
								firstTimeOfRest?.Name?.FormatEx("-", false)));
							typeAndTypeRest.Add(new Tuple<string, string>(" ",
								secondTimeOfRest?.Name?.FormatEx("-", false)));
						}
						else
						{
							typeAndTypeRest.Add(new Tuple<string, string>("Время отдыха",
								request.NullSafe(r => r.TimeOfRest.Name.FormatEx("-", false)) +
								(request.Tour != null
									? $" {request.Tour.DateIncome.FormatEx()}-{request.Tour.DateOutcome.FormatEx()}"
									: string.Empty)));
						}

						if (request.SubjectOfRest != null)
						{
							typeAndTypeRest.Add(new Tuple<string, string>("Тематика смены",
								request.SubjectOfRest.Name.FormatEx("-", false)));
						}

						if (request.TransferTo != null)
						{
							typeAndTypeRest.Add(new Tuple<string, string>(
								"Проезд из города Москвы к месту отдыха",
								request.TransferTo.Name.FormatEx("-", false)));
						}

						if (request.TransferFrom != null)
						{
							typeAndTypeRest.Add(new Tuple<string, string>(
								"Проезд из места отдыха в город Москву",
								request.TransferFrom.Name.FormatEx("-", false)));
						}
					}
					else
					{
						typeAndTypeRest.Add(new Tuple<string, string>("Год кампании",
							request.YearOfRest?.Name.FormatEx("-", false)));
					}

					AppendBlock(doc, "Цель обращения", typeAndTypeRest);

					if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation &&
						request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest &&
						request.TypeOfRestId != (long)TypeOfRestEnum.MoneyOn3To7 &&
						request.TypeOfRestId != (long)TypeOfRestEnum.MoneyOn7To15)
					{
						if (!request.IsFirstCompany)
						{
							AppendBlock(doc, "Место отдыха", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Регион",
									request.NullSafe(r => r.PlaceOfRest.Name.FormatEx("-", false))),
								new Tuple<string, string>("Организация отдыха и оздоровления",
									request.NullSafe(r => r.Tour.Hotels.Name.FormatEx("-", false)) +
									$" ({request.NullSafe(r => r.Tour.Hotels.Address.FormatEx("-", false))})")
							});
						}
						else
						{
							var firstPlaceOfRest = request.PlacesOfRest?.Where(ss => ss.Order > 0)
								.OrderBy(ss => ss.Order).Select(ss => ss.PlaceOfRest).FirstOrDefault();
							var secondPlaceOfRest = request.PlacesOfRest?.Where(ss => ss.Order > 0)
								.OrderBy(ss => ss.Order).Select(ss => ss.PlaceOfRest).LastOrDefault();
							if (secondPlaceOfRest?.Id == firstPlaceOfRest?.Id)
							{
								secondPlaceOfRest = null;
							}

							AppendBlock(doc, "Место отдыха", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Приоритетный регион отдыха",
									request.NullSafe(r => r.PlaceOfRest.Name.FormatEx("-", false))),
								new Tuple<string, string>("Дополнительные регионы отдыха",
									firstPlaceOfRest?.Name?.FormatEx("-", false)),
								new Tuple<string, string>(" ", secondPlaceOfRest?.Name?.FormatEx("-", false))
							});
						}

						if (isNotYouthRest)
						{
							var placement = new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Детей",
									request.NullSafe(r => r.CountPlace.FormatEx("-", false)))
							};

							if (request.CountAttendants > 0)
							{
								placement.Add(new Tuple<string, string>("Сопровождающих",
									request.NullSafe(r => r.CountAttendants.FormatEx("-", false))));
							}

							if (request.BookingGuid.HasValue)
							{
								if (bookings.Any())
								{
									var locations =
										bookings.Select(
											b =>
												new LocationRequest
												{
													Count = b.CountRooms,
													Name = b.TourVolume.TypeOfRooms.Name,
													RoomId = b.TourVolume.TypeOfRoomsId ?? 0
												}).ToList();

									var placementText = new StringBuilder();

									foreach (var location in locations)
									{
										placementText.AppendFormat("{0} (количество номеров: {1})\n", location.Name,
											location.Count);
									}

									placement.Add(new Tuple<string, string>("Размещение", placementText.ToString()));
								}
							}

							AppendBlock(doc, "Размещение", placement);
						}
						else
						{
							AppendBlock(doc, "Размещение", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Отдыхающих", "1")
							});
						}
					}
					else if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation
							 && request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest &&
							 !string.IsNullOrWhiteSpace(request.BankName))
					{
						AppendBlock(doc, "Банковские реквизиты", new List<Tuple<string, string>>
						{
							new Tuple<string, string>("Наименование банка", request.BankName.FormatEx("-", false)),
							new Tuple<string, string>("ИНН банка", request.BankInn.FormatEx("-", false)),
							new Tuple<string, string>("КПП банка", request.BankKpp.FormatEx("-", false)),
							new Tuple<string, string>("БИК", request.BankBik.FormatEx("-", false)),
							new Tuple<string, string>("Расчетный счет", request.BankAccount.FormatEx("-", false)),
							new Tuple<string, string>("Корреспондентский счёт", request.BankCorr.FormatEx("-", false)),
							new Tuple<string, string>("Номер карты", request.BankCardNumber.FormatEx("-", false)),
							new Tuple<string, string>("Фамилия получателя", request.BankLastName.FormatEx("-", false)),
							new Tuple<string, string>("Имя получателя", request.BankFirstName.FormatEx("-", false)),
							new Tuple<string, string>("Отчество получателя",
								request.BankMiddleName.FormatEx("-", false))
						});
					}
					else
					{
						AppendBlock(doc, "Путевки", null);

						foreach (var iv in request.InformationVouchers)
						{
							var tuples = new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Цель обращения",
									iv.NullSafe(r => r.Type.Name.FormatEx("-", false))),
								new Tuple<string, string>("Наименование организации",
									iv.NullSafe(r => r.OrganizationName.FormatEx("-", false))),
								new Tuple<string, string>("Дата начала", iv?.DateFrom.FormatEx()),
								new Tuple<string, string>("Дата окончания", iv?.DateTo.FormatEx()),
								new Tuple<string, string>("Стоимость (руб.)", iv?.Price.FormatEx()),
								new Tuple<string, string>("Стоимость дороги(руб)", iv?.CostOfRide.FormatEx()),
								new Tuple<string, string>("Количество отдохнувших", iv?.CountPeople.FormatEx())
							};

							foreach (var ap in iv?.AttendantsPrice ?? new List<RequestInformationVoucherAttendant>())
							{
								tuples.Add(new Tuple<string, string>(
									$"{ap?.Applicant?.LastName} {ap?.Applicant?.FirstName} {ap?.Applicant?.MiddleName}",
									$"Стоимость путевки: {ap?.Price.FormatEx().Trim()};\nСтоимость дороги: {ap?.CostOfRide.FormatEx().Trim()}"));
							}

							AppendBlock(doc, "", tuples);
						}

						if (!string.IsNullOrWhiteSpace(request.BankName))
						{
							AppendBlock(doc, "Банковские реквизиты", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Наименование банка", request.BankName.FormatEx("-", false)),
								new Tuple<string, string>("ИНН банка", request.BankInn.FormatEx("-", false)),
								new Tuple<string, string>("КПП банка", request.BankKpp.FormatEx("-", false)),
								new Tuple<string, string>("БИК", request.BankBik.FormatEx("-", false)),
								new Tuple<string, string>("Лицевой счет", request.BankAccount.FormatEx("-", false)),
								new Tuple<string, string>("Номер карты", request.BankCardNumber.FormatEx("-", false))
							});
						}
					}

					AppendBlock(doc, "Сведения о заявителе", new List<Tuple<string, string>>
					{
						new Tuple<string, string>("Фамилия",
							request.NullSafe(r => r.Applicant.LastName.FormatEx("-", false))),
						new Tuple<string, string>("Имя",
							request.NullSafe(r => r.Applicant.FirstName.FormatEx("-", false))),
						new Tuple<string, string>("Отчество",
							request.NullSafe(r => r.Applicant.MiddleName.FormatEx("-", false))),
						new Tuple<string, string>("Пол", request?.Applicant?.Male ?? false ? "Мужской" : "Женский"),
						new Tuple<string, string>("Дата рождения",
							request.NullSafe(r => r.Applicant.DateOfBirth.FormatEx("-", false))),
						new Tuple<string, string>("Место рождения",
							request.NullSafe(r => r.Applicant.PlaceOfBirth.FormatEx("-", false))),
						new Tuple<string, string>("СНИЛС",
							request.NullSafe(r => r.Applicant.Snils.FormatEx("-", false))),
						new Tuple<string, string>("Телефон",
							request.NullSafe(r => r.Applicant.Phone.FormatEx("-", false))),
						new Tuple<string, string>("Email",
							request.NullSafe(r => r.Applicant.Email.FormatEx("-", false))),
						new Tuple<string, string>("Тип документа, удостоверяющего личность",
							request.NullSafe(r => r.Applicant.DocumentType.Name.FormatEx("-", false))),
						new Tuple<string, string>("Серия и номер",
							request.NullSafe(
								r => r.Applicant.DocumentSeria.FormatEx("-", false) + " " +
									 r.Applicant.DocumentNumber.FormatEx("-", false))),
						new Tuple<string, string>("Когда выдан документ",
							request.NullSafe(r => r.Applicant.DocumentDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
						new Tuple<string, string>("Кем выдан документ",
							request.NullSafe(r => r.Applicant.DocumentSubjectIssue.FormatEx("-", false))),
						new Tuple<string, string>("Заявитель является сопровождающим",
							request.NullSafe(r => r.Applicant.IsAccomp) ? "Да" : "Нет"),
						new Tuple<string, string>("Заявление подается представителем заявителя",
							request.NullSafe(r => r.AgentApplicant ?? false) ? "Да" : "Нет")
					});

					if (request.AgentApplicant ?? false)
					{
						var attendantAgent = request?.Attendant?.FirstOrDefault(a => a.IsAgent);

						var block = new List<Tuple<string, string>>
						{
							new Tuple<string, string>("Фамилия",
								request.NullSafe(r => r.Agent.LastName.FormatEx("-", false))),
							new Tuple<string, string>("Имя",
								request.NullSafe(r => r.Agent.FirstName.FormatEx("-", false))),
							new Tuple<string, string>("Отчество",
								request.NullSafe(r => r.Agent.MiddleName.FormatEx("-", false))),
							new Tuple<string, string>("Пол", request?.Agent?.Male ?? false ? "Мужской" : "Женский"),
							new Tuple<string, string>("Дата рождения",
								request.NullSafe(r => r.Agent.DateOfBirth.FormatEx("-", false))),
							new Tuple<string, string>("Место рождения",
								request.NullSafe(r => r.Agent.PlaceOfBirth.FormatEx("-", false))),
							new Tuple<string, string>("Телефон",
								request.NullSafe(r => r.Agent.Phone.FormatEx("-", false))),
							new Tuple<string, string>("Email",
								request.NullSafe(r => r.Agent.Email.FormatEx("-", false))),
							new Tuple<string, string>("Тип документа, удостоверяющего личность",
								request.NullSafe(r => r.Agent.DocumentType.Name.FormatEx("-", false))),
							new Tuple<string, string>("Серия и номер",
								request.NullSafe(
									r => r.Agent.DocumentSeria.FormatEx("-", false) + " " +
										 r.Agent.DocumentNumber.FormatEx("-", false))),
							new Tuple<string, string>("Когда выдан документ",
								request.NullSafe(r => r.Agent.DocumentDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
							new Tuple<string, string>("Кем выдан документ",
								request.NullSafe(r => r.Agent.DocumentSubjectIssue.FormatEx("-", false))),
							new Tuple<string, string>("Представитель заявителя является сопровождающим",
								(attendantAgent != null).FormatEx())
						};

						if (request.StatusApplicant == "4" && request.Child != null && request.Child.Any())
						{
							var name = request.RepresentInterest?.Name ?? request.Agent?.StatusByChild?.Name;

							if (string.IsNullOrWhiteSpace(name))
							{
								if (request.IsFirstCompany && (
									request.TypeOfRestId == (long)TypeOfRestEnum.ChildRest ||
									request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParents ||
									request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsPoor ||
									request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsComplex ||
									request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalid ||
									request.TimeOfRestId == (long)TypeOfRestEnum.RestWithInvalid4To17 ||
									request.TimeOfRestId == (long)TypeOfRestEnum.RestWithWithrestrictionsInvalid4To17 ||
									request.TypeOfRestId ==
									(long)TypeOfRestEnum.RestWithParentsInvalidOrphanComplex ||
									request.TypeOfRestId == (long)TypeOfRestEnum.ChildRestCamps ||
									request.TypeOfRestId == (long)TypeOfRestEnum.TentChildrenCamp ||
									request.TypeOfRestId == (long)TypeOfRestEnum.ChildRestFederalCamps))
								{
									name = request?.Applicant?.Male == true ? "Отец" : "Мать";
								}
							}

							if (!string.IsNullOrWhiteSpace(name))
							{
								block.Insert(0, new Tuple<string, string>("Представляет интересы", name));
							}
						}

						AppendBlock(doc, "Сведения о представителе заявителя", block);

						AppendBlock(doc, "Сведения о доверенности на подачу заявления", new List<Tuple<string, string>>
						{
							new Tuple<string, string>("Дата выдачи доверенности",
								request.NullSafe(r => r.Agent.ProxyDateOfIssure.FormatEx("dd.MM.yyyy", "-"))),
							new Tuple<string, string>("Дата окончания срока действия доверенности",
								request.NullSafe(r => r.Agent.ProxyEndDate.FormatEx("dd.MM.yyyy", "-"))),
							new Tuple<string, string>("ФИО нотариуса",
								request.NullSafe(r => r.Agent.NotaryName.FormatEx("-", false))),
							new Tuple<string, string>("Номер доверенности",
								request.NullSafe(r => r.Agent.ProxyNumber.FormatEx("-", false)))
						});

						if (attendantAgent != null)
						{
							AppendBlock(doc, "Сведения о доверенности на сопровождение", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Дата выдачи доверенности",
									attendantAgent.NullSafe(r => r.ProxyDateOfIssure.FormatEx("-", false))),
								new Tuple<string, string>("Дата окончания срока действия доверенности",
									attendantAgent.NullSafe(r => r.ProxyEndDate.FormatEx("-", false))),
								new Tuple<string, string>("ФИО нотариуса",
									attendantAgent.NullSafe(r => r.NotaryName.FormatEx("-", false))),
								new Tuple<string, string>("Номер доверенности",
									attendantAgent.NullSafe(r => r.ProxyNumber.FormatEx("-", false)))
							});
						}
					}

					if (request.Child != null && request.Child.Any(c => !c.IsDeleted))
					{
						AppendTitle(doc, "Сведения о детях");
						foreach (var child in request.Child.Where(c => !c.IsDeleted).ToList())
						{
							AppendBlock(doc, "", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Фамилия",
									child.NullSafe(r => r.LastName.FormatEx("-", false))),
								new Tuple<string, string>("Имя", child.NullSafe(r => r.FirstName.FormatEx("-", false))),
								new Tuple<string, string>("Отчество",
									child.NullSafe(r => r.MiddleName.FormatEx("-", false))),
								new Tuple<string, string>("Пол", child.NullSafe(r => r.Male) ? "Мужской" : "Женский"),
								new Tuple<string, string>("Дата рождения",
									child.NullSafe(r => r.DateOfBirth.FormatEx("dd.MM.yyyy", "-"))),
								new Tuple<string, string>("Место рождения",
									child.NullSafe(r => r.PlaceOfBirth.FormatEx("-", false))),
								new Tuple<string, string>("Тип документа, удостоверяющего личность",
									child.NullSafe(r => r.DocumentType.Name.FormatEx("-", false))),
								new Tuple<string, string>("Серия и номер",
									child.NullSafe(r =>
										r.DocumentSeria.FormatEx("-", false) + " " +
										r.DocumentNumber.FormatEx("-", false))),
								new Tuple<string, string>("Когда выдан документ",
									child.NullSafe(r => r.DocumentDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
								new Tuple<string, string>("Кем выдан документ",
									child.NullSafe(r => r.DocumentSubjectIssue.FormatEx("-", false))),
								new Tuple<string, string>("СНИЛС", child.NullSafe(r => r.Snils.FormatEx("-", false)))
							});
							if (child.BenefitType != null)
							{
								AppendBlock(doc, "Сведения о категории льготы ребенка",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Категория",
											child.NullSafe(r => r.BenefitType.Name.FormatEx("-", false))),
										new Tuple<string, string>("Вид ограничения",
											child.NullSafe(r => r.TypeOfRestriction.Name.FormatEx("-", false))),
										new Tuple<string, string>("Подвид ограничения",
											child.NullSafe(r => r.TypeOfSubRestriction.Name.FormatEx("-", false))),
										new Tuple<string, string>("",
											child.NullSafe(r => r.BenefitAnswerComment.FormatEx("-", false)))
									}.Where(t => t.Item2 != "-").ToList());

								if (child.BenefitType.NeedApproveDocument)
								{
									AppendBlock(doc,
										"Документ, подтверждающий, что ребенок находится в трудной жизненной ситуации",
										new List<Tuple<string, string>>
										{
											new Tuple<string, string>("Кем выдан документ",
												child.NullSafe(r => r.BenefitSubjectIssue.FormatEx("-", false))),
											new Tuple<string, string>("Когда выдан документ",
												child.NullSafe(r => r.BenefitDateOfIssure.FormatEx("dd.MM.yyyy", "-"))),
											new Tuple<string, string>("Номер документа",
												child.NullSafe(r => r.BenefitNumber.FormatEx("-", false)))
										});
								}
							}

							if (!request.IsFirstCompany)
							{
								if (child.School != null && !child.SchoolNotPresent)
								{
									AppendBlock(doc, "Образовательное учреждение", new List<Tuple<string, string>>
									{
										new Tuple<string, string>("",
											child.NullSafe(r => r.School.Name.FormatEx("-", false)))
									});
								}

								if (child.SchoolNotPresent)
								{
									AppendBlock(doc, "Образовательное учреждение", new List<Tuple<string, string>>
									{
										new Tuple<string, string>("", "Нет в списке")
									});
								}
							}

							if (request.NullSafe(r => r.PlaceOfRest.IsForegin))
							{
								AppendBlock(doc, "Документ удостоверяющий личность ребенка за рубежом",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Фамилия",
											child.NullSafe(r => r.ForeginLastName.FormatEx("-", false))),
										new Tuple<string, string>("Имя",
											child.NullSafe(r => r.ForeginName.FormatEx("-", false))),
										new Tuple<string, string>(
											"Тип документа, удостоверяющего личность ребенка за рубежом",
											child.NullSafe(r => r.ForeginType.Name.FormatEx("-", false))),
										new Tuple<string, string>("Серия и номер",
											child.NullSafe(r =>
												r.ForeginSeria.FormatEx("-", false) + " " +
												r.ForeginNumber.FormatEx("-", false))),
										new Tuple<string, string>("Дата выдачи",
											child.NullSafe(r => r.ForeginDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
										new Tuple<string, string>("Срок действия",
											child.NullSafe(r => r.ForeginDateEnd.FormatEx("dd.MM.yyyy", "-"))),
										new Tuple<string, string>("Кем выдан",
											child.NullSafe(r => r.ForeginSubjectIssue.FormatEx("-", false)))
									});
							}

							if (child?.Address?.BtiAddress != null)
							{
								AppendBlock(
									doc,
									"Адрес регистрации",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Адрес",
											child.Address.ToString().FormatEx("-", false))
									});
							}
							else if (!string.IsNullOrWhiteSpace(child?.Address?.Appartment)
									 || !string.IsNullOrWhiteSpace(child?.Address?.Street)
									 || !string.IsNullOrWhiteSpace(child?.Address?.House)
									 || !string.IsNullOrWhiteSpace(child?.Address?.Corpus)
									 || !string.IsNullOrWhiteSpace(child?.Address?.Vladenie)
									 || !string.IsNullOrWhiteSpace(child?.Address?.Stroenie))
							{
								AppendBlock(
									doc,
									"Адрес регистрации",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Адрес",
											(child.Address?.ToString()).FormatEx("-", false))
									});
							}

							if (child.ApplicantId.HasValue)
							{
								AppendBlock(
									doc,
									"Сопровождающий",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("ФИО",
											$"{child.Applicant.LastName.FormatEx("-", false)} {child.Applicant.FirstName.FormatEx("-", false)} {child.Applicant.MiddleName.FormatEx("-", false)}"),
										new Tuple<string, string>("Статус по отношению к ребенку",
											(child.StatusByChild?.Name).FormatEx("-", false))
									});
							}

							if (request.TypeOfRestId == (long)TypeOfRestEnum.Compensation ||
								request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest)
							{
								AppendBlock(
									doc,
									"Путевка",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Путевка",
											$"{child.RequestInformationVoucher?.OrganizationName.FormatEx("-", false) ?? "-"} - {child.RequestInformationVoucher?.DateFrom.FormatEx("-", false) ?? "-"}"),
										new Tuple<string, string>("Стоимость путевки (руб)",
											child.CostOfTour?.FormatEx() ?? "-"),
										new Tuple<string, string>("Стоимость проезда (руб)",
											child.CostOfRide?.FormatEx() ?? "-")
									});
							}
						}
					}

					if (request.Attendant != null && request.Attendant.Any(a => !a.IsAgent))
					{
						AppendBlock(doc, "Сведения о сопровождающих", null);
						foreach (var attendant in request.Attendant.Where(a => !a.IsAgent && !a.IsDeleted)?.ToList() ??
												  new List<Applicant>())
						{
							var block = new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Является доверенным лицом", attendant.IsProxy.FormatEx()),
								new Tuple<string, string>("Фамилия",
									attendant.NullSafe(r => r.LastName.FormatEx("-", false))),
								new Tuple<string, string>("Имя",
									attendant.NullSafe(r => r.FirstName.FormatEx("-", false))),
								new Tuple<string, string>("Отчество",
									attendant.NullSafe(r => r.MiddleName.FormatEx("-", false))),
								new Tuple<string, string>("Пол", attendant?.Male ?? false ? "Мужской" : "Женский"),
								new Tuple<string, string>("СНИЛС", (attendant?.Snils).FormatEx("-", false)),
								new Tuple<string, string>("Телефон",
									attendant.NullSafe(r => r.Phone.FormatEx("-", false))),
								new Tuple<string, string>("Дата рождения",
									attendant.NullSafe(r => r.DateOfBirth.FormatEx("dd.MM.yyyy", "-"))),
								new Tuple<string, string>("Место рождения",
									attendant.NullSafe(r => r.PlaceOfBirth.FormatEx("-", false))),
								new Tuple<string, string>("Email",
									attendant.NullSafe(r => r.Email.FormatEx("-", false))),
								new Tuple<string, string>("Тип документа, удостоверяющего личность",
									attendant.NullSafe(r => r.DocumentType.Name.FormatEx("-", false))),
								new Tuple<string, string>("Серия и номер",
									attendant.NullSafe(r =>
										r.DocumentSeria.FormatEx("-", false) + " " +
										r.DocumentNumber.FormatEx("-", false))),
								new Tuple<string, string>("Когда выдан документ",
									attendant.NullSafe(r => r.DocumentDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
								new Tuple<string, string>("Кем выдан документ",
									attendant.NullSafe(r => r.DocumentSubjectIssue.FormatEx("-", false))),
								new Tuple<string, string>("Статус по отношению к ребёнку",
									attendant.ApplicantType?.Name?.FormatEx("-", false))
							};

							AppendBlock(doc, "", block);
							if (attendant.IsProxy)
							{
								AppendBlock(doc, "Сведения о доверенности на сопровождение",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Дата выдачи доверенности",
											attendant.NullSafe(r => r.ProxyDateOfIssure.FormatEx("-", false))),
										new Tuple<string, string>("Дата окончания срока действия доверенности",
											attendant.NullSafe(r => r.ProxyEndDate.FormatEx("-", false))),
										new Tuple<string, string>("ФИО нотариуса",
											attendant.NullSafe(r => r.NotaryName.FormatEx("-", false))),
										new Tuple<string, string>("Номер доверенности",
											attendant.NullSafe(r => r.ProxyNumber.FormatEx("-", false)))
									});
							}
						}
					}

					AppendFooterNotificationOfRegistration(doc);
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = (request.RequestNumber ?? request.Id.ToString(CultureInfo.InvariantCulture)) + "" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		/// Отказ (Из реестра отказов) в документе
		/// </summary>
		/// <param name="request"></param>
		/// <param name="unitOfWork"></param>
		/// <returns></returns>
		public IDocument DeclineDocument(Request request, ICollection<Booking> bookings)
		{

			if (request == null || request.IsDeleted) return null;

			var isNotYouthRest = request.TypeOfRestId != (long)TypeOfRestEnum.YouthRestOrphanCamps &&
								 request.TypeOfRestId != (long)TypeOfRestEnum.YouthRestCamps;

			MemoryStream ms;

			using (ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					mainPart.Document = doc;

					var section = new SectionProperties();
					section.AppendChild(new PageMargin { Bottom = 1000, Top = 1000, Left = 1000, Right = 1000 });
					mainPart.Document.Body.AppendChild(section);

					AppendHeader(doc);
					var generalBlock = new List<Tuple<string, string>>
					{
						new Tuple<string, string>("Номер отказа", request.RequestNumber.FormatEx()),
						new Tuple<string, string>("Дата отказа", request.DateRequest.FormatEx("dd.MM.yyyy")),
						new Tuple<string, string>("Источник отказа", request.Source.Name.FormatEx("-", false))
					};

					if (!string.IsNullOrEmpty(request.RequestNumberMpgu))
					{
						generalBlock.Add(new Tuple<string, string>("Номер заявления МПГУ", request.RequestNumberMpgu));
					}

					if (request.DeclineReason != null)
					{
						generalBlock.Add(new Tuple<string, string>("Причина отказа", request.DeclineReason.Name));
					}

					if (!string.IsNullOrWhiteSpace(request.StatusApplicant) && isNotYouthRest)
					{
						var statusApplicant = GetStatusApplicantName(request.StatusApplicant);

						if (!string.IsNullOrWhiteSpace(statusApplicant))
						{
							generalBlock.Add(new Tuple<string, string>("Заявление подает", statusApplicant));
						}
					}

					AppendBlock(doc, "Общие сведения об отказе", generalBlock);

					var typeAndTypeRest = new List<Tuple<string, string>>
					{
						new Tuple<string, string>("Цель обращения",
							request.NullSafe(r => r.TypeOfRest.Name.FormatEx("-", false)))
					};

					if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation &&
						request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest &&
						request.TypeOfRestId != (long)TypeOfRestEnum.MoneyOn3To7 &&
						request.TypeOfRestId != (long)TypeOfRestEnum.MoneyOn7To15)
					{
						if (request.IsFirstCompany)
						{
							var firstTimeOfRest = request.TimesOfRest?.Where(ss => ss.Order > 0)
								.Select(ss => ss.TimeOfRest).FirstOrDefault();
							var secondTimeOfRest = request.TimesOfRest?.Where(ss => ss.Order > 0)
								.Select(ss => ss.TimeOfRest).LastOrDefault();
							if (secondTimeOfRest?.Id == firstTimeOfRest?.Id)
							{
								secondTimeOfRest = null;
							}

							typeAndTypeRest.Add(new Tuple<string, string>("Приоритетное время отдыха",
								request.TimeOfRest?.Name.FormatEx("-", false)));
							typeAndTypeRest.Add(new Tuple<string, string>("Дополнительные времена отдыха",
								firstTimeOfRest?.Name?.FormatEx("-", false)));
							typeAndTypeRest.Add(new Tuple<string, string>(" ",
								secondTimeOfRest?.Name?.FormatEx("-", false)));
						}
						else
						{
							typeAndTypeRest.Add(new Tuple<string, string>("Время отдыха",
								request.NullSafe(r => r.TimeOfRest.Name.FormatEx("-", false)) +
								(request.Tour != null
									? $" {request.Tour.DateIncome.FormatEx()}-{request.Tour.DateOutcome.FormatEx()}"
									: string.Empty)));
						}

						if (request.SubjectOfRest != null)
						{
							typeAndTypeRest.Add(new Tuple<string, string>("Тематика смены",
								request.SubjectOfRest.Name.FormatEx("-", false)));
						}

						if (request.TransferTo != null)
						{
							typeAndTypeRest.Add(new Tuple<string, string>(
								"Проезд из города Москвы к месту отдыха",
								request.TransferTo.Name.FormatEx("-", false)));
						}

						if (request.TransferFrom != null)
						{
							typeAndTypeRest.Add(new Tuple<string, string>(
								"Проезд из места отдыха в город Москву",
								request.TransferFrom.Name.FormatEx("-", false)));
						}
					}
					else
					{
						typeAndTypeRest.Add(new Tuple<string, string>("Год кампании",
							request.YearOfRest?.Name.FormatEx("-", false)));
					}

					AppendBlock(doc, "Цель обращения", typeAndTypeRest);

					if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation &&
						request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest &&
						request.TypeOfRestId != (long)TypeOfRestEnum.MoneyOn3To7 &&
						request.TypeOfRestId != (long)TypeOfRestEnum.MoneyOn7To15)
					{
						if (!request.IsFirstCompany)
						{
							AppendBlock(doc, "Место отдыха", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Регион",
									request.NullSafe(r => r.PlaceOfRest.Name.FormatEx("-", false))),
								new Tuple<string, string>("Организация отдыха и оздоровления",
									request.NullSafe(r => r.Tour.Hotels.Name.FormatEx("-", false)) +
									$" ({request.NullSafe(r => r.Tour.Hotels.Address.FormatEx("-", false))})")
							});
						}
						else
						{
							var firstPlaceOfRest = request.PlacesOfRest?.Where(ss => ss.Order > 0)
								.OrderBy(ss => ss.Order).Select(ss => ss.PlaceOfRest).FirstOrDefault();
							var secondPlaceOfRest = request.PlacesOfRest?.Where(ss => ss.Order > 0)
								.OrderBy(ss => ss.Order).Select(ss => ss.PlaceOfRest).LastOrDefault();
							if (secondPlaceOfRest?.Id == firstPlaceOfRest?.Id)
							{
								secondPlaceOfRest = null;
							}

							AppendBlock(doc, "Место отдыха", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Приоритетный регион отдыха",
									request.NullSafe(r => r.PlaceOfRest.Name.FormatEx("-", false))),
								new Tuple<string, string>("Дополнительные регионы отдыха",
									firstPlaceOfRest?.Name?.FormatEx("-", false)),
								new Tuple<string, string>(" ", secondPlaceOfRest?.Name?.FormatEx("-", false))
							});
						}

						if (isNotYouthRest)
						{
							var placement = new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Детей",
									request.NullSafe(r => r.CountPlace.FormatEx("-", false)))
							};

							if (request.CountAttendants > 0)
							{
								placement.Add(new Tuple<string, string>("Сопровождающих",
									request.NullSafe(r => r.CountAttendants.FormatEx("-", false))));
							}

							if (request.BookingGuid.HasValue)
							{
								if (bookings.Any())
								{
									var locations =
										bookings.Select(
											b =>
												new LocationRequest
												{
													Count = b.CountRooms,
													Name = b.TourVolume.TypeOfRooms.Name,
													RoomId = b.TourVolume.TypeOfRoomsId ?? 0
												}).ToList();

									var placementText = new StringBuilder();

									foreach (var location in locations)
									{
										placementText.AppendFormat("{0} (количество номеров: {1})\n", location.Name,
											location.Count);
									}

									placement.Add(new Tuple<string, string>("Размещение", placementText.ToString()));
								}
							}

							AppendBlock(doc, "Размещение", placement);
						}
						else
						{
							AppendBlock(doc, "Размещение", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Отдыхающих", "1")
							});
						}
					}
					else if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation
							 && request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest &&
							 !string.IsNullOrWhiteSpace(request.BankName))
					{
						AppendBlock(doc, "Банковские реквизиты", new List<Tuple<string, string>>
						{
							new Tuple<string, string>("Наименование банка", request.BankName.FormatEx("-", false)),
							new Tuple<string, string>("ИНН банка", request.BankInn.FormatEx("-", false)),
							new Tuple<string, string>("КПП банка", request.BankKpp.FormatEx("-", false)),
							new Tuple<string, string>("БИК", request.BankBik.FormatEx("-", false)),
							new Tuple<string, string>("Расчетный счет", request.BankAccount.FormatEx("-", false)),
							new Tuple<string, string>("Корреспондентский счёт", request.BankCorr.FormatEx("-", false)),
							new Tuple<string, string>("Номер карты", request.BankCardNumber.FormatEx("-", false)),
							new Tuple<string, string>("Фамилия получателя", request.BankLastName.FormatEx("-", false)),
							new Tuple<string, string>("Имя получателя", request.BankFirstName.FormatEx("-", false)),
							new Tuple<string, string>("Отчество получателя",
								request.BankMiddleName.FormatEx("-", false))
						});
					}
					else
					{
						AppendBlock(doc, "Путевки", null);

						foreach (var iv in request.InformationVouchers)
						{
							var tuples = new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Цель обращения",
									iv.NullSafe(r => r.Type.Name.FormatEx("-", false))),
								new Tuple<string, string>("Наименование организации",
									iv.NullSafe(r => r.OrganizationName.FormatEx("-", false))),
								new Tuple<string, string>("Дата начала", iv?.DateFrom.FormatEx()),
								new Tuple<string, string>("Дата окончания", iv?.DateTo.FormatEx()),
								new Tuple<string, string>("Стоимость (руб.)", iv?.Price.FormatEx()),
								new Tuple<string, string>("Стоимость дороги(руб)", iv?.CostOfRide.FormatEx()),
								new Tuple<string, string>("Количество отдохнувших", iv?.CountPeople.FormatEx())
							};

							foreach (var ap in iv?.AttendantsPrice ?? new List<RequestInformationVoucherAttendant>())
							{
								tuples.Add(new Tuple<string, string>(
									$"{ap?.Applicant?.LastName} {ap?.Applicant?.FirstName} {ap?.Applicant?.MiddleName}",
									$"Стоимость путевки: {ap?.Price.FormatEx().Trim()};\nСтоимость дороги: {ap?.CostOfRide.FormatEx().Trim()}"));
							}

							AppendBlock(doc, "", tuples);
						}

						if (!string.IsNullOrWhiteSpace(request.BankName))
						{
							AppendBlock(doc, "Банковские реквизиты", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Наименование банка", request.BankName.FormatEx("-", false)),
								new Tuple<string, string>("ИНН банка", request.BankInn.FormatEx("-", false)),
								new Tuple<string, string>("КПП банка", request.BankKpp.FormatEx("-", false)),
								new Tuple<string, string>("БИК", request.BankBik.FormatEx("-", false)),
								new Tuple<string, string>("Лицевой счет", request.BankAccount.FormatEx("-", false)),
								new Tuple<string, string>("Номер карты", request.BankCardNumber.FormatEx("-", false))
							});
						}
					}

					AppendBlock(doc, "Сведения о заявителе", new List<Tuple<string, string>>
					{
						new Tuple<string, string>("Фамилия",
							request.NullSafe(r => r.Applicant.LastName.FormatEx("-", false))),
						new Tuple<string, string>("Имя",
							request.NullSafe(r => r.Applicant.FirstName.FormatEx("-", false))),
						new Tuple<string, string>("Отчество",
							request.NullSafe(r => r.Applicant.MiddleName.FormatEx("-", false))),
						new Tuple<string, string>("Пол", request?.Applicant?.Male ?? false ? "Мужской" : "Женский"),
						new Tuple<string, string>("Дата рождения",
							request.NullSafe(r => r.Applicant.DateOfBirth.FormatEx("-", false))),
						new Tuple<string, string>("Место рождения",
							request.NullSafe(r => r.Applicant.PlaceOfBirth.FormatEx("-", false))),
						new Tuple<string, string>("СНИЛС",
							request.NullSafe(r => r.Applicant.Snils.FormatEx("-", false))),
						new Tuple<string, string>("Телефон",
							request.NullSafe(r => r.Applicant.Phone.FormatEx("-", false))),
						new Tuple<string, string>("Email",
							request.NullSafe(r => r.Applicant.Email.FormatEx("-", false))),
						new Tuple<string, string>("Тип документа, удостоверяющего личность",
							request.NullSafe(r => r.Applicant.DocumentType.Name.FormatEx("-", false))),
						new Tuple<string, string>("Серия и номер",
							request.NullSafe(
								r => r.Applicant.DocumentSeria.FormatEx("-", false) + " " +
									 r.Applicant.DocumentNumber.FormatEx("-", false))),
						new Tuple<string, string>("Когда выдан документ",
							request.NullSafe(r => r.Applicant.DocumentDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
						new Tuple<string, string>("Кем выдан документ",
							request.NullSafe(r => r.Applicant.DocumentSubjectIssue.FormatEx("-", false))),
						new Tuple<string, string>("Заявитель является сопровождающим",
							request.NullSafe(r => r.Applicant.IsAccomp) ? "Да" : "Нет"),
						new Tuple<string, string>("Заявление подается представителем заявителя",
							request.NullSafe(r => r.AgentApplicant ?? false) ? "Да" : "Нет")
					});

					if (request.AgentApplicant ?? false)
					{
						var attendantAgent = request?.Attendant?.FirstOrDefault(a => a.IsAgent);

						var block = new List<Tuple<string, string>>
						{
							new Tuple<string, string>("Фамилия",
								request.NullSafe(r => r.Agent.LastName.FormatEx("-", false))),
							new Tuple<string, string>("Имя",
								request.NullSafe(r => r.Agent.FirstName.FormatEx("-", false))),
							new Tuple<string, string>("Отчество",
								request.NullSafe(r => r.Agent.MiddleName.FormatEx("-", false))),
							new Tuple<string, string>("Пол", request?.Agent?.Male ?? false ? "Мужской" : "Женский"),
							new Tuple<string, string>("Дата рождения",
								request.NullSafe(r => r.Agent.DateOfBirth.FormatEx("-", false))),
							new Tuple<string, string>("Место рождения",
								request.NullSafe(r => r.Agent.PlaceOfBirth.FormatEx("-", false))),
							new Tuple<string, string>("Телефон",
								request.NullSafe(r => r.Agent.Phone.FormatEx("-", false))),
							new Tuple<string, string>("Email",
								request.NullSafe(r => r.Agent.Email.FormatEx("-", false))),
							new Tuple<string, string>("Тип документа, удостоверяющего личность",
								request.NullSafe(r => r.Agent.DocumentType.Name.FormatEx("-", false))),
							new Tuple<string, string>("Серия и номер",
								request.NullSafe(
									r => r.Agent.DocumentSeria.FormatEx("-", false) + " " +
										 r.Agent.DocumentNumber.FormatEx("-", false))),
							new Tuple<string, string>("Когда выдан документ",
								request.NullSafe(r => r.Agent.DocumentDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
							new Tuple<string, string>("Кем выдан документ",
								request.NullSafe(r => r.Agent.DocumentSubjectIssue.FormatEx("-", false))),
							new Tuple<string, string>("Представитель заявителя является сопровождающим",
								(attendantAgent != null).FormatEx())
						};

						if (request.StatusApplicant == "4" && request.Child != null && request.Child.Any())
						{
							var name = request.RepresentInterest?.Name ?? request.Agent?.StatusByChild?.Name;

							if (string.IsNullOrWhiteSpace(name))
							{
								if (request.IsFirstCompany && (
									request.TypeOfRestId == (long)TypeOfRestEnum.ChildRest ||
									request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParents ||
									request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsPoor ||
									request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsComplex ||
									request.TypeOfRestId == (long)TypeOfRestEnum.RestWithParentsInvalid ||
									request.TypeOfRestId ==
									(long)TypeOfRestEnum.RestWithParentsInvalidOrphanComplex ||
									request.TypeOfRestId == (long)TypeOfRestEnum.ChildRestCamps ||
									request.TypeOfRestId == (long)TypeOfRestEnum.TentChildrenCamp ||
									request.TypeOfRestId == (long)TypeOfRestEnum.ChildRestFederalCamps))
								{
									name = request?.Applicant?.Male == true ? "Отец" : "Мать";
								}
							}

							if (!string.IsNullOrWhiteSpace(name))
							{
								block.Insert(0, new Tuple<string, string>("Представляет интересы", name));
							}
						}

						AppendBlock(doc, "Сведения о представителе заявителя", block);

						AppendBlock(doc, "Сведения о доверенности на подачу заявления", new List<Tuple<string, string>>
						{
							new Tuple<string, string>("Дата выдачи доверенности",
								request.NullSafe(r => r.Agent.ProxyDateOfIssure.FormatEx("dd.MM.yyyy", "-"))),
							new Tuple<string, string>("Дата окончания срока действия доверенности",
								request.NullSafe(r => r.Agent.ProxyEndDate.FormatEx("dd.MM.yyyy", "-"))),
							new Tuple<string, string>("ФИО нотариуса",
								request.NullSafe(r => r.Agent.NotaryName.FormatEx("-", false))),
							new Tuple<string, string>("Номер доверенности",
								request.NullSafe(r => r.Agent.ProxyNumber.FormatEx("-", false)))
						});

						if (attendantAgent != null)
						{
							AppendBlock(doc, "Сведения о доверенности на сопровождение", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Дата выдачи доверенности",
									attendantAgent.NullSafe(r => r.ProxyDateOfIssure.FormatEx("-", false))),
								new Tuple<string, string>("Дата окончания срока действия доверенности",
									attendantAgent.NullSafe(r => r.ProxyEndDate.FormatEx("-", false))),
								new Tuple<string, string>("ФИО нотариуса",
									attendantAgent.NullSafe(r => r.NotaryName.FormatEx("-", false))),
								new Tuple<string, string>("Номер доверенности",
									attendantAgent.NullSafe(r => r.ProxyNumber.FormatEx("-", false)))
							});
						}
					}

					if (request.Child != null && request.Child.Any(c => !c.IsDeleted))
					{
						AppendTitle(doc, "Сведения о детях");
						foreach (var child in request.Child.Where(c => !c.IsDeleted).ToList())
						{
							AppendBlock(doc, "", new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Фамилия",
									child.NullSafe(r => r.LastName.FormatEx("-", false))),
								new Tuple<string, string>("Имя", child.NullSafe(r => r.FirstName.FormatEx("-", false))),
								new Tuple<string, string>("Отчество",
									child.NullSafe(r => r.MiddleName.FormatEx("-", false))),
								new Tuple<string, string>("Пол", child.NullSafe(r => r.Male) ? "Мужской" : "Женский"),
								new Tuple<string, string>("Дата рождения",
									child.NullSafe(r => r.DateOfBirth.FormatEx("dd.MM.yyyy", "-"))),
								new Tuple<string, string>("Место рождения",
									child.NullSafe(r => r.PlaceOfBirth.FormatEx("-", false))),
								new Tuple<string, string>("Тип документа, удостоверяющего личность",
									child.NullSafe(r => r.DocumentType.Name.FormatEx("-", false))),
								new Tuple<string, string>("Серия и номер",
									child.NullSafe(r =>
										r.DocumentSeria.FormatEx("-", false) + " " +
										r.DocumentNumber.FormatEx("-", false))),
								new Tuple<string, string>("Когда выдан документ",
									child.NullSafe(r => r.DocumentDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
								new Tuple<string, string>("Кем выдан документ",
									child.NullSafe(r => r.DocumentSubjectIssue.FormatEx("-", false))),
								new Tuple<string, string>("СНИЛС", child.NullSafe(r => r.Snils.FormatEx("-", false)))
							});
							if (child.BenefitType != null)
							{
								AppendBlock(doc, "Сведения о категории льготы ребенка",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Категория",
											child.NullSafe(r => r.BenefitType.Name.FormatEx("-", false))),
										new Tuple<string, string>("Вид ограничения",
											child.NullSafe(r => r.TypeOfRestriction.Name.FormatEx("-", false))),
										new Tuple<string, string>("Подвид ограничения",
											child.NullSafe(r => r.TypeOfSubRestriction.Name.FormatEx("-", false))),
										new Tuple<string, string>("",
											child.NullSafe(r => r.BenefitAnswerComment.FormatEx("-", false)))
									}.Where(t => t.Item2 != "-").ToList());

								if (child.BenefitType.NeedApproveDocument)
								{
									AppendBlock(doc,
										"Документ, подтверждающий, что ребенок находится в трудной жизненной ситуации",
										new List<Tuple<string, string>>
										{
											new Tuple<string, string>("Кем выдан документ",
												child.NullSafe(r => r.BenefitSubjectIssue.FormatEx("-", false))),
											new Tuple<string, string>("Когда выдан документ",
												child.NullSafe(r => r.BenefitDateOfIssure.FormatEx("dd.MM.yyyy", "-"))),
											new Tuple<string, string>("Номер документа",
												child.NullSafe(r => r.BenefitNumber.FormatEx("-", false)))
										});
								}
							}

							if (!request.IsFirstCompany)
							{
								if (child.School != null && !child.SchoolNotPresent)
								{
									AppendBlock(doc, "Образовательное учреждение", new List<Tuple<string, string>>
									{
										new Tuple<string, string>("",
											child.NullSafe(r => r.School.Name.FormatEx("-", false)))
									});
								}

								if (child.SchoolNotPresent)
								{
									AppendBlock(doc, "Образовательное учреждение", new List<Tuple<string, string>>
									{
										new Tuple<string, string>("", "Нет в списке")
									});
								}
							}

							if (request.NullSafe(r => r.PlaceOfRest.IsForegin))
							{
								AppendBlock(doc, "Документ удостоверяющий личность ребенка за рубежом",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Фамилия",
											child.NullSafe(r => r.ForeginLastName.FormatEx("-", false))),
										new Tuple<string, string>("Имя",
											child.NullSafe(r => r.ForeginName.FormatEx("-", false))),
										new Tuple<string, string>(
											"Тип документа, удостоверяющего личность ребенка за рубежом",
											child.NullSafe(r => r.ForeginType.Name.FormatEx("-", false))),
										new Tuple<string, string>("Серия и номер",
											child.NullSafe(r =>
												r.ForeginSeria.FormatEx("-", false) + " " +
												r.ForeginNumber.FormatEx("-", false))),
										new Tuple<string, string>("Дата выдачи",
											child.NullSafe(r => r.ForeginDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
										new Tuple<string, string>("Срок действия",
											child.NullSafe(r => r.ForeginDateEnd.FormatEx("dd.MM.yyyy", "-"))),
										new Tuple<string, string>("Кем выдан",
											child.NullSafe(r => r.ForeginSubjectIssue.FormatEx("-", false)))
									});
							}

							if (child?.Address?.BtiAddress != null)
							{
								AppendBlock(
									doc,
									"Адрес регистрации",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Адрес",
											child.Address.ToString().FormatEx("-", false))
									});
							}
							else if (!string.IsNullOrWhiteSpace(child?.Address?.Appartment)
									 || !string.IsNullOrWhiteSpace(child?.Address?.Street)
									 || !string.IsNullOrWhiteSpace(child?.Address?.House)
									 || !string.IsNullOrWhiteSpace(child?.Address?.Corpus)
									 || !string.IsNullOrWhiteSpace(child?.Address?.Vladenie)
									 || !string.IsNullOrWhiteSpace(child?.Address?.Stroenie))
							{
								AppendBlock(
									doc,
									"Адрес регистрации",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Адрес",
											(child.Address?.ToString()).FormatEx("-", false))
									});
							}

							if (child.ApplicantId.HasValue)
							{
								AppendBlock(
									doc,
									"Сопровождающий",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("ФИО",
											$"{child.Applicant.LastName.FormatEx("-", false)} {child.Applicant.FirstName.FormatEx("-", false)} {child.Applicant.MiddleName.FormatEx("-", false)}"),
										new Tuple<string, string>("Статус по отношению к ребенку",
											(child.StatusByChild?.Name).FormatEx("-", false))
									});
							}

							if (request.TypeOfRestId == (long)TypeOfRestEnum.Compensation ||
								request.TypeOfRestId == (long)TypeOfRestEnum.CompensationYouthRest)
							{
								AppendBlock(
									doc,
									"Путевка",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Путевка",
											$"{child.RequestInformationVoucher?.OrganizationName.FormatEx("-", false) ?? "-"} - {child.RequestInformationVoucher?.DateFrom.FormatEx("-", false) ?? "-"}"),
										new Tuple<string, string>("Стоимость путевки (руб)",
											child.CostOfTour?.FormatEx() ?? "-"),
										new Tuple<string, string>("Стоимость проезда (руб)",
											child.CostOfRide?.FormatEx() ?? "-")
									});
							}
						}
					}

					if (request.Attendant != null && request.Attendant.Any(a => !a.IsAgent))
					{
						AppendBlock(doc, "Сведения о сопровождающих", null);
						foreach (var attendant in request.Attendant.Where(a => !a.IsAgent && !a.IsDeleted)?.ToList() ??
												  new List<Applicant>())
						{
							var block = new List<Tuple<string, string>>
							{
								new Tuple<string, string>("Является доверенным лицом", attendant.IsProxy.FormatEx()),
								new Tuple<string, string>("Фамилия",
									attendant.NullSafe(r => r.LastName.FormatEx("-", false))),
								new Tuple<string, string>("Имя",
									attendant.NullSafe(r => r.FirstName.FormatEx("-", false))),
								new Tuple<string, string>("Отчество",
									attendant.NullSafe(r => r.MiddleName.FormatEx("-", false))),
								new Tuple<string, string>("Пол", attendant?.Male ?? false ? "Мужской" : "Женский"),
								new Tuple<string, string>("СНИЛС", (attendant?.Snils).FormatEx("-", false)),
								new Tuple<string, string>("Телефон",
									attendant.NullSafe(r => r.Phone.FormatEx("-", false))),
								new Tuple<string, string>("Дата рождения",
									attendant.NullSafe(r => r.DateOfBirth.FormatEx("dd.MM.yyyy", "-"))),
								new Tuple<string, string>("Место рождения",
									attendant.NullSafe(r => r.PlaceOfBirth.FormatEx("-", false))),
								new Tuple<string, string>("Email",
									attendant.NullSafe(r => r.Email.FormatEx("-", false))),
								new Tuple<string, string>("Тип документа, удостоверяющего личность",
									attendant.NullSafe(r => r.DocumentType.Name.FormatEx("-", false))),
								new Tuple<string, string>("Серия и номер",
									attendant.NullSafe(r =>
										r.DocumentSeria.FormatEx("-", false) + " " +
										r.DocumentNumber.FormatEx("-", false))),
								new Tuple<string, string>("Когда выдан документ",
									attendant.NullSafe(r => r.DocumentDateOfIssue.FormatEx("dd.MM.yyyy", "-"))),
								new Tuple<string, string>("Кем выдан документ",
									attendant.NullSafe(r => r.DocumentSubjectIssue.FormatEx("-", false))),
								new Tuple<string, string>("Статус по отношению к ребёнку",
									attendant.ApplicantType?.Name?.FormatEx("-", false))
							};

							AppendBlock(doc, "", block);
							if (attendant.IsProxy)
							{
								AppendBlock(doc, "Сведения о доверенности на сопровождение",
									new List<Tuple<string, string>>
									{
										new Tuple<string, string>("Дата выдачи доверенности",
											attendant.NullSafe(r => r.ProxyDateOfIssure.FormatEx("-", false))),
										new Tuple<string, string>("Дата окончания срока действия доверенности",
											attendant.NullSafe(r => r.ProxyEndDate.FormatEx("-", false))),
										new Tuple<string, string>("ФИО нотариуса",
											attendant.NullSafe(r => r.NotaryName.FormatEx("-", false))),
										new Tuple<string, string>("Номер доверенности",
											attendant.NullSafe(r => r.ProxyNumber.FormatEx("-", false)))
									});
							}
						}
					}

					AppendFooterNotificationOfRegistration(doc);
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = (request.RequestNumber ?? request.Id.ToString(CultureInfo.InvariantCulture)) + "" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		//странный метод childId на вход есть, но его не использует
		/// <summary>
		///     Получение печатной формы для межведомственного запроса
		/// </summary>
		public IDocument InteragencyRequestDocument(Child child, Account account)
		{
			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					var mainRunProperties = new RunProperties();
					mainRunProperties.AppendChild(new RunFonts
					{
						Ascii = "Times New Roman",
						HighAnsi = "Times New Roman",
						ComplexScript = "Times New Roman"
					});
					mainRunProperties.AppendChild(new FontSize { Val = "24" });

					var mainTitleRunProperties = mainRunProperties.CloneNode(true);
					mainTitleRunProperties.AppendChild(new Bold());

					var mainSmallRunProperties = mainRunProperties.CloneNode(true);
					mainSmallRunProperties.RemoveAllChildren<FontSize>();
					mainSmallRunProperties.AppendChild(new FontSize { Val = "20" });

					var mainUnderlineRunProperties = mainRunProperties.CloneNode(true);

					var bottomBorders = new ParagraphBorders
					{
						BottomBorder =
							new BottomBorder
							{
								Val = new EnumValue<BorderValues>(BorderValues.Single),
								Size = 6,
								Space = 1
							}
					};

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center },
							new SpacingBetweenLines { After = "20" }),
						new Run(mainTitleRunProperties.CloneNode(true),
							new Text("МЕЖВЕДОМСТВЕННЫЙ ЗАПРОС"))));
					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center },
							new SpacingBetweenLines { After = "20" }),
						new Run(mainTitleRunProperties.CloneNode(true),
							new Text("О ПРЕДСТАВЛЕНИИ ИНФОРМАЦИИ, "))));
					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center },
							new SpacingBetweenLines { After = "20" }),
						new Run(mainTitleRunProperties.CloneNode(true),
							new Text("НЕОБХОДИМОЙ ДЛЯ ПРЕДОСТАВЛЕНИЯ ГОСУДАРСТВЕННОЙ УСЛУГИ"))));

					doc.AppendChild(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "20" })));
					doc.AppendChild(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "20" })));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left },
							bottomBorders.CloneNode(true),
							new SpacingBetweenLines { After = "20" }),
						new Run(mainRunProperties.CloneNode(true),
							new Text(
								"В целях предоставления государственной услуги по организации отдыха и оздоровления детей в соответствии с постановлением Правительства Москвы от 15 февраля 2011 г. № 29-ПП \"Об организации отдыха и оздоровления детей города Москвы в 2011 году и последующие годы\" просим представить в отношении"))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left },
							bottomBorders.CloneNode(true),
							new SpacingBetweenLines { After = "20" }),
						new Run(mainUnderlineRunProperties.CloneNode(true), new Text(""))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center },
							new SpacingBetweenLines { After = "10" }),
						new Run(mainSmallRunProperties.CloneNode(true),
							new Text("(Ф.И.О. лица, подающего заявление, его местонахождение"))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center },
							new SpacingBetweenLines { After = "10" }),
						new Run(mainSmallRunProperties.CloneNode(true),
							new Text(
								"(для юридического лица)/место жительства (для физического лица) либо иные сведения, необходимые "))));

					doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(mainSmallRunProperties.CloneNode(true),
							new Text("для представления информации)"))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left },
							bottomBorders.CloneNode(true),
							new SpacingBetweenLines { After = "20" }),
						new Run(mainRunProperties.CloneNode(true),
							new Text(
								"следующую информацию об отнесении к льготной категории \"дети из малообеспеченной семьи\" ребенка заявителя"))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left },
							bottomBorders.CloneNode(true),
							new SpacingBetweenLines { After = "20" }),
						new Run(mainUnderlineRunProperties.CloneNode(true), new Text(""))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center },
							new SpacingBetweenLines { After = "10" }),
						new Run(mainSmallRunProperties.CloneNode(true),
							new Text("(Ф.И.О. лица, подающего заявление, его местонахождение"))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center },
							new SpacingBetweenLines { After = "10" }),
						new Run(mainSmallRunProperties.CloneNode(true),
							new Text(
								"(для юридического лица)/место жительства (для физического лица) либо иные сведения, необходимые "))));

					doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(mainSmallRunProperties.CloneNode(true),
							new Text("для представления информации)"))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left },
							new SpacingBetweenLines { After = "20" }),
						new Run(mainRunProperties.CloneNode(true),
							new Text("Ответ на межведомственный запрос просим направить по"))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left },
							bottomBorders.CloneNode(true),
							new SpacingBetweenLines { After = "20" }),
						new Run(mainUnderlineRunProperties.CloneNode(true), new Text(""))));

					doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(mainSmallRunProperties.CloneNode(true),
							new Text(
								"(указать адрес электронной почты Исполнителя, номер факса, по которым надлежит направить ответ)"))));

					doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
						new Run(mainRunProperties.CloneNode(true),
							new Text("в срок до ______________________________."))));

					doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
						new Run(mainSmallRunProperties.CloneNode(true),
							new Text("(указать срок ожидаемого ответа)"))));

					mainRunProperties = new RunProperties();
					mainRunProperties.AppendChild(new RunFonts
					{
						Ascii = "Times New Roman",
						HighAnsi = "Times New Roman",
						ComplexScript = "Times New Roman"
					});
					mainRunProperties.AppendChild(new FontSize { Val = "24" });

					var captionRunProperties = new RunProperties();
					captionRunProperties.AppendChild(new RunFonts
					{
						Ascii = "Times New Roman",
						HighAnsi = "Times New Roman",
						ComplexScript = "Times New Roman"
					});
					captionRunProperties.AppendChild(new FontSize { Val = "18" });

					var tblProp = new TableProperties(
						new TableBorders(
							new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
							new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
							new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
							new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
							new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
							new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));

					doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
						new Run(mainRunProperties.CloneNode(true),
							new Text("Уведомление выдал:"))));

					var table = new Table();
					table.AppendChild(tblProp.CloneNode(true));

					var row = new TableRow();

					var cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" },
							new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) })),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(mainRunProperties.CloneNode(true),
							new Text(account.Position.FormatEx()))));
					row.AppendChild(cell);


					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "250" }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(new Run(new Text(" "))));
					row.AppendChild(cell);

					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1604" },
							new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) })),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(new Run(new Text(" "))));
					row.AppendChild(cell);

					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "250" }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(new Run(new Text(" "))));
					row.AppendChild(cell);

					cell = new TableCell();
					cell.AppendChild(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" },
						new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom }
					));

					cell.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(mainRunProperties.CloneNode(true),
							new Text(account.Name.FormatEx()))));
					row.AppendChild(cell);

					table.AppendChild(row);
					row = new TableRow();

					cell = new TableCell();
					cell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(captionRunProperties.CloneNode(true),
							new Text("(должность)"))));
					row.AppendChild(cell);

					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "250" }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(new Run(new Text(" "))));
					row.AppendChild(cell);

					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1604" }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(captionRunProperties.CloneNode(true), new Text("(подпись)"))));
					row.AppendChild(cell);

					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "250" }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(new Run(new Text(" "))));
					row.AppendChild(cell);

					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(captionRunProperties.CloneNode(true), new Text("(расшифровка подписи)"))));
					row.AppendChild(cell);

					table.AppendChild(row);

					doc.AppendChild(table);

					mainPart.Document = doc;
				}
				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Межведомственный запрос" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		///     уведомление об отсутствии мест.
		/// </summary>
		public IDocument NotificationEmpty(TypeOfRest typeOfRest, Domain.TimeOfRest timeOfRest, int countChildren,
			int countAttendant, PlaceOfRest placeOfRest, Account account)
		{
			MemoryStream ms;
			using (ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());
					DocumentHeader(doc, "__________________");

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize("28").Bold(),
							new Text(
								"Уведомление об отсутствии в наличии путевки на отдых и оздоровление в организации отдыха по выбранным заявителем параметрам"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(" " + DateTime.Now.FormatEx()))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true), new Text("Данные заявителя: ")),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("____________________________________________________________"))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true), new Text("Контактная информация: ")),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("___________________________ __________________________"))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Indentation { FirstLine = "4000" },
								new SpacingBetweenLines { Line = "276" },
								new Justification { Val = JustificationValues.Left }),
							new Run(new RunProperties().SetFont().SetFontSize("16"),
								new Text(
									"телефон                                                       (электронная почта) "))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Наименование запрашиваемой услуги:  ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("__________________________________________________"))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Цель обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true), new Text(typeOfRest?.Name))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Год кампании: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true), new Text("2016"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Число детей льготных категорий: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(countChildren.FormatEx()))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Число сопровождающих: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(countAttendant.FormatEx()))));

					if (timeOfRest != null)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = "20" }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Время отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true), new Text(timeOfRest.Name))));
					}

					if (placeOfRest != null)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = "20" }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Место отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true), new Text(placeOfRest.Name))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения обращения: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(((RunProperties)titleRequestRunPropertiesItalic.CloneNode(true)).Underline(),
								new Text(
									"Отсутствие в наличии путевки на отдых и оздоровление в организации отдыха по выбранным заявителем параметрам."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: пункт 6 ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"приложения к приказу Департамента культуры города Москвы от 3 марта 2016 г.  № 130 \"Об организации в 2016 году выездного отдыха и оздоровления детей льготных категорий, имеющих место жительства в города Москве\" – заявление подается на предоставление путевки на отдых и оздоровление из числа имеющихся в наличии путевок на отдых и оздоровление, закупленных для соответствующей категории и вида отдыха."))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					SignBlockNotification(doc, account, " ");
					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отсутствии места" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		///     Уведомление об отсутствии возможности замены сертификата на отдых и оздоровление на  бесплатную путевку для отдыха
		///     и оздоровления (в случае выбора сертификата на отдых и оздоровление)
		/// </summary>
		public IDocument NotificationLackOfPossibilityReplacement(Request request, Account account)
		{
			MemoryStream ms;
			using (ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());


					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize("28").Bold(),
							new Text(
								"Уведомление об отсутствии возможности замены сертификата на отдых и оздоровление на  бесплатную путевку для отдыха и оздоровления (в случае выбора сертификата на отдых и оздоровление)"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();
					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = "20" },
								new Indentation { FirstLine = "600" }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
										"В соответствии с постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" в случае выбора сертификата на отдых и оздоровление (далее – сертификат) замена на бесплатную путевку для отдыха и оздоровления (далее – бесплатная путевка) ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("не допускается.")
								{ Space = SpaceProcessingModeValues.Preserve })));

					var applicant = request?.Applicant ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new Indentation { FirstLine = "500" },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									$"Я, {applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, подтверждаю, что ознакомлен(а) с отсутствием возможности замены сертификата на бесплатную путевку.")
								{
									Space = SpaceProcessingModeValues.Preserve
								})));

					SignBlockNotification2020(doc, account, "Исполнитель:");
					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об отсутствии возможности замены" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension,
				};
			}
		}

		/// <summary>
		///     Уведомление об ознакомлении с перечнем противопоказаний к отдыху и оздоровлению, утвержденным Министерством
		///     здравоохранения Российской Федерации.
		/// </summary>
		public IDocument NotificationAcquaintance(Request request, Account account)
		{
			MemoryStream ms;
			using (ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize("28").Bold(),
							new Text(
								"Уведомление об ознакомлении с перечнем противопоказаний к отдыху и оздоровлению, утвержденным Министерством здравоохранения Российской Федерации"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var body = new[]
					{
						"В соответствии с приказом Министерства здравоохранения Российской Федерации от 13 июня 2018 г. № 327н \"Об утверждении порядка оказания медицинской помощи несовершеннолетним в период оздоровления и организованного отдыха\" в организации отдыха и оздоровления (далее – детские лагеря) направляются дети, не имеющие следующих медицинских противопоказаний для пребывания в детских лагерях:",
						"соматические заболевания в острой и подострой стадии, хронические заболевания в стадии обострения, в стадии декомпенсации;",
						"инфекционные и паразитарные болезни, в том числе с поражением глаз и кожи, инфестации (педикулез, чесотка) – в период до окончания срока изоляции;",
						"установленный диагноз \"бактерионосительство возбудителей кишечных инфекций, дифтерии\";",
						"активный туберкулез любой локализации;",
						"наличие контакта с инфекционными больными в течение 21 календарного дня перед заездом в детский лагерь;",
						"отсутствие профилактических прививок в случае возникновения массовых инфекционных заболеваний или при угрозе возникновения эпидемий;",
						"злокачественные новообразования, требующие лечения, в том числе проведения химиотерапии;",
						"эпилепсия с текущими приступами, в том числе резистентная к проводимому лечению;",
						"эпилепсия с медикаментозной ремиссией менее 1 года;",
						"кахексия;",
						"психические расстройства и расстройства поведения, вызванные употреблением психоактивных веществ, а также иные психические расстройства и расстройства поведения в состоянии обострения и (или) представляющие опасность для ребенка и окружающих;",
						"хронические заболевания, требующие соблюдения назначенного лечащим врачом режима лечения (диета, прием лекарственных препаратов для медицинского применения и специализированных продуктов лечебного питания) (для детских лагерей палаточного типа);",
					};

					foreach (var s in body)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new Indentation { FirstLine = "500" },
									new SpacingBetweenLines { After = "20" }),
								new Run(titleRequestRunProperties.CloneNode(true),
									new Text(s)
									{
										Space = SpaceProcessingModeValues.Preserve
									})));
					}

					var applicant = request?.Applicant ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new Indentation { FirstLine = "500" },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text(
									$"Я, {applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, подтверждаю, что ознакомлен(а) с перечнем противопоказаний к отдыху и оздоровлению, утвержденным Министерством здравоохранения Российской Федерации.")
								{
									Space = SpaceProcessingModeValues.Preserve
								})));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					SignBlockNotification2020(doc, account, "Исполнитель:");
					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление об ознакомлении с перечнем" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension,
				};
			}
		}

		/// <summary>
		///     Уведомление о непризнании отказа от осуществления выездного отдыха отказом по уважительным причинам
		/// </summary>
		public IDocument NotificationNonRecognition(Request request, Account account)
		{
			MemoryStream ms;
			using (ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize("28").Bold(),
							new Text(
								"Уведомление о непризнании отказа от осуществления выездного отдыха отказом по уважительным причинам"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);

					var applicant = request.NullSafe(r => r.Applicant) ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));
					doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = "20" }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер путевки: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.CertificateNumber))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Организация отдыха и оздоровления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.Tour?.Hotels?.Name))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Время отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text($"{request.TimeOfRest?.Name} ({request.Tour?.DateIncome.FormatEx()} - {request.Tour?.DateOutcome.FormatEx()})"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание приёма заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunProperties.CloneNode(true),
								new Text("Порядок организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации, утвержденный постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\" (далее - Порядок)."))));

					doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = "20" }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text("Ваше заявление об отказе от предоставленной бесплатной путевки для отдыха и оздоровления по уважительной причине ее неиспользования(далее – заявление) рассмотрено."))));

					var strings = new[]
					{
						"В соответствии с пунктом 10.2 Порядка отказ от осуществления выездного отдыха на основании предоставленной бесплатной путевки для отдыха и оздоровления (далее – бесплатная путевка) подтверждается документами о наличии уважительных причин.",
						"К уважительным причинам неосуществления отдыха относятся:",
						"заболевание, травма ребенка, лица из числа детей-сирот и детей, оставшихся без попечения родителей, на дату заезда в организацию отдыха и оздоровления ",
						"заболевание, травма, смерть сопровождающего лица (в случае организации совместного выездного отдыха) на дату заезда в организацию отдыха и оздоровления; ",
						"необходимость осуществления сопровождающим лицом ухода за больным членом семьи (в случае организации совместного выездного отдыха) на дату заезда в организацию отдыха и оздоровления; ",
						"карантин ребенка, карантин лица из числа детей-сирот и детей, оставшихся без попечения родителей, карантин лица, проживающего совместно с ребенком, лицом из числа детей-сирот и детей, оставшихся без попечения родителей, а также в случае организации совместного выездного отдыха карантин сопровождающего лица на дату заезда в организацию отдыха и оздоровления; ",
						"смерть близкого родственника ребенка (родителя, бабушки, дедушки, брата, сестры) в период с даты окончания заявочной кампании в целях организации отдыха и оздоровления до даты заезда в организацию отдыха и оздоровления; ",
						"получение детьми-инвалидами, детьми с ограниченными возможностями здоровья санаторно-курортного лечения или реабилитации в то же время, на которое предоставлена бесплатная путевка для отдыха и оздоровления; ",
						"выявление у ребенка, сопровождающего лица в период с даты окончания заявочной кампании в целях организации отдыха и оздоровления медицинских противопоказаний для пребывания в организации отдыха и оздоровления, смены климатической зоны, а также способа проезда в организацию отдыха и оздоровления.",
						"Прилагаемый к Вашему заявлению документ не является документом, подтверждающим уважительную причину неиспользования предоставленной бесплатной путевки.",
						$"Учитывая изложенное, в соответствии с пунктом 10.4.2 Порядка оформить отказ от бесплатной путевки № { request.CertificateNumber} в установленном порядке не представляется возможным."

					};

					foreach (var s in strings)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = "20" }, new Indentation { FirstLine = "500" }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true), new Text(s))));
					}

					SignSimpleBlockNotification(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о непризнании отказа" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension,
				};
			}
		}

		/// <summary>
		///     Уведомление о непризнании отказа от осуществления выездного отдыха отказом по уважительным причинам (в связи с
		///     подачей заявления на отказ после установленного срока)
		/// </summary>
		public IDocument NotificationNonRecognitionByTime(Request request, Account account)
		{
			MemoryStream ms;
			using (ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize("28").Bold(),
							new Text(
								"Уведомление о непризнании отказа от осуществления выездного отдыха отказом по уважительным причинам (в связи с подачей заявления на отказ после установленного срока)"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);

					var applicant = request.NullSafe(r => r.Applicant) ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation &&
						request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = "20" }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.RequestNumber.FormatEx()))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatEx(string.Empty)}"))));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					var strings = new[]
					{
						" ",
						"Ваше заявление об отказе от предоставленной бесплатной путевки для отдыха и оздоровления рассмотрено.",
						"В соответствии с постановлением Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\", отказ от осуществления выездного отдыха на основании предоставленной бесплатной путевки для отдыха и оздоровления подтверждается документами о наличии уважительных причин.",
						"К уважительным причинам относятся:",
						"заболевание, травма ребенка, лица из числа детей - сирот и детей, оставшихся без попечения родителей;",
						"заболевание, травма сопровождающего лица (в случае организации совместного выездного отдыха);",
						"необходимость осуществления сопровождающим лицом ухода за больным членом семьи (в случае организации совместного выездного отдыха);",
						"карантин ребенка, карантин лица из числа детей-сирот и детей, оставшихся без попечения родителей, карантин лица, проживающего совместно с ребенком, лицом из числа детей-сирот и детей, оставшихся без попечения родителей, а также в случае организации совместного выездного отдыха карантин сопровождающего лица;",
						"смерть близкого родственника (родителя, бабушки, дедушки, брата, сестры, дяди, тети).",
						$"Учитывая, что заявление об отказе от предоставленной путевки для отдыха и оздоровления № {request.CertificateNumber} было подано по истечении 60 календарных дней со дня начала периода отдыха и оздоровления, указанного в бесплатной путевке для отдыха и оздоровления, оформить отказ от осуществления выездного отдыха на основании бесплатной путевки для отдыха и оздоровления в установленном порядке не представляется возможным.",
						" ",
						" ",
						" "
					};

					foreach (var s in strings)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = "20" }, new Indentation { FirstLine = "500" }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true), new Text(s))));
					}

					SignSimpleBlockNotification(doc, account);
					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о непризнании отказа" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension,
				};
			}
		}

		/// <summary>
		///     Уведомление о непринятии отказа от рассмотрения поданного заявления на предоставление бесплатной путевки для отдыха
		///     и оздоровления.
		/// </summary>
		public IDocument NotificationRefuseRefuse(Request request, Account account)
		{
			MemoryStream ms;
			using (ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(new RunProperties().SetFont().SetFontSize("28").Bold(),
							new Text(
								"Уведомление о непринятии заявления об отзыве заявления о предоставлении услуг отдыха и оздоровления"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(new RunProperties().SetFont().SetFontSize("28").Bold(), new Text(" "))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.NullSafe(r => r.Applicant) ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					if (request.TypeOfRestId != (long)TypeOfRestEnum.Compensation ||
						request.TypeOfRestId != (long)TypeOfRestEnum.CompensationYouthRest)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = "20" }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(request.RequestNumber.FormatEx()))));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatEx(string.Empty)}"))));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Причина обращения: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"прекращение рассмотрения поданного заявления о предоставлении услуг отдыха и оздоровления."))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Результат рассмотрения заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"отказ в приеме заявления об отзыве заявления о предоставлении услуг отдыха и оздоровления."))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = "20" }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Основание: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									"постановление Правительства Москвы от 22 февраля 2017 г. № 56-ПП \"Об организации отдыха и оздоровления детей, находящихся в трудной жизненной ситуации\""))));

					SignBlockNotification(doc, account, $"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}");
					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о непринятии отказа" + Extension,
					ContentType = _WORDPROCESSINGML_CONTENT_TYPE,
					MimeType = MimeType,
					MimeTypeShort = Extension,
				};
			}
		}

		#region PrivateMethods

		private void SignSimpleBlockNotification(Document doc, Account account, string sign = "Уведомление выдал:")
		{
			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = "24" });

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));
			var captionRunProperties = new RunProperties().SetFont().SetFontSize("16");

			var table = new Table();
			table.AppendChild(tblProp.CloneNode(true));

			var row = new TableRow();

			var cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text("Уведомление выдал:"))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.AppendChild(new TableCellProperties(
				new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "5717" },
				new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom }
			));
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(account.Name.FormatEx() +
							 (string.IsNullOrWhiteSpace(account.Position) ? "" : $", {account.Position}")))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(" "))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.AppendChild(new TableCellProperties(
				new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1804" },
				new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom }));
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(DateTime.Now.Date.FormatEx()))));
			row.AppendChild(cell);

			table.AppendChild(row);
			// -----------------------------------------------------------
			row = new TableRow();

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(" "))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "5717" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(captionRunProperties.CloneNode(true), new Text("(ФИО работника, должность)"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(" "))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1804" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true), captionRunProperties.CloneNode(true),
					new Text("(дата)"))));
			row.AppendChild(cell);

			table.AppendChild(row);

			doc.AppendChild(table);
		}

		private void DocumentHeader(Document doc, string requestNumber)
		{
			var titleProp = new RunProperties().SetFont().SetFontSize("28").Bold();
			doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleProp.CloneNode(true), new Text("ДЕПАРТАМЕНТ КУЛЬТУРЫ ГОРОДА МОСКВЫ"))));
			doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleProp.CloneNode(true),
					new Text("Государственное автономное учреждение культуры города Москвы"))));
			doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleProp.CloneNode(true), new Text("\"Московское агентство организации отдыха и туризма\""))));
			var pp = new ParagraphProperties(new Justification { Val = JustificationValues.Center });
			pp.AppendChild(new ParagraphBorders().BottomBorder("000000", 3));
			doc.AppendChild(new Paragraph(pp, new Run(titleProp.CloneNode(true), new Text("(ГАУК «МОСГОРТУР»)"))));
			var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

			doc.AppendChild(
				new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Left },
						new SpacingBetweenLines { After = "20" }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text("№ " + requestNumber))));
		}

		private void SignWorkerCompensationBlock(Document doc, Account account = null)
		{
			account = account ?? new Account();

			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = _SIZE_24 });

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));
			var captionRunProperties = new RunProperties().SetFont().SetFontSize(_SIZE_18);

			var table = new Table();
			table.AppendChild(tblProp.CloneNode(true));

			var row = new TableRow();

			var cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1231" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text("Уведомление выдал:"))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.AppendChild(new TableCellProperties(
				new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "6931" },
				new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom }
			));
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(
						$"{account.Name.FormatEx()}{(string.IsNullOrWhiteSpace(account.Position) ? string.Empty : $", {account.Position}")}"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(" "))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.AppendChild(new TableCellProperties(
				new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" },
				new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom }));

			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true), new Text(DateTime.Today.FormatEx()))));
			row.AppendChild(cell);

			table.AppendChild(row);
			// -----------------------------------------------------------
			row = new TableRow();

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1231" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(" "))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "6931" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(captionRunProperties.CloneNode(true), new Text("(Ф.И.О. работника, должность)"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(" "))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true), captionRunProperties.CloneNode(true),
					new Text("(дата)"))));
			row.AppendChild(cell);


			table.AppendChild(row);

			doc.AppendChild(table);
		}

		private void AppendChildrenToDocument(Document doc, Request request)
		{
			var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

			var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
			titleRequestRunPropertiesBold.AppendChild(new Bold());
			var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);

			if (request.Child != null)
			{
				foreach (var child in request.Child.Where(c => !c.IsDeleted))
				{
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text(
										"Данные ребёнка/лица из числа детей-сирот и детей, оставшихся без попечения родителей: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Льготная категория: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text($"{child.BenefitType?.Name}"))));
				}
			}

			if (request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps)
			{
				var child = request.Applicant;
				doc.AppendChild(
					new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left },
							new SpacingBetweenLines { After = _SIZE_20 }),
						new Run(titleRequestRunPropertiesBold.CloneNode(true),
							new Text("Данные о ребенке: ") { Space = SpaceProcessingModeValues.Preserve }),
						new Run(titleRequestRunPropertiesItalic.CloneNode(true),
							new Text(
								$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatEx(string.Empty)}"))));

				doc.AppendChild(
					new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left },
							new SpacingBetweenLines { After = _SIZE_20 }),
						new Run(titleRequestRunPropertiesBold.CloneNode(true),
							new Text("Льготная категория: ") { Space = SpaceProcessingModeValues.Preserve }),
						new Run(titleRequestRunPropertiesItalic.CloneNode(true),
							new Text(
								"Лица, из числа детей-сирот и детей, оставшихся без попечения родителей, в возрасте от 18 до 23 лет (включительно), обучающиеся по образовательным программам среднего профессионального образования или образовательным программам высшего образования по очной форме обучения"))));
			}
		}

		/// <summary>
		/// Подпись работника для компенсации
		/// </summary>
		private void SignWorkerBlock(Document doc, Account account, string name = "Принял:")
		{
			//заглушка
			account = account ?? new Account();
			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = "22" });

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));

			var captionRunProperties = new RunProperties().SetFont().SetFontSizeSupperscript();



			var table = new Table();
			table.AppendChild(tblProp.CloneNode(true));

			//отступ - нулевая строка
			{
				doc.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Both },
				new SpacingBetweenLines { After = _SIZE_20 })));
			}

			var row = new TableRow();

			var cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1231" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(name))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.Append(new TableCellProperties(
				new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "6931" },
				new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom }
			));
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(
						$"{account?.Name?.FormatEx()}{(string.IsNullOrWhiteSpace(account?.Position) ? string.Empty : $", {account.Position}")}"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" },
					new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) })),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			table.AppendChild(row);
			// -----------------------------------------------------------
			row = new TableRow();

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1231" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(_SPACE))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "6931" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(captionRunProperties.CloneNode(true), new Text("(Ф.И.О. работника, должность)"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true), captionRunProperties.CloneNode(true),
					new Text("(подпись работника)"))));
			row.AppendChild(cell);


			table.AppendChild(row);

			doc.AppendChild(table);
		}

		/// <summary>
		/// Формирование таблицы об услугах, оказанных ребенку\детям в течение последних 3-х лет
		/// </summary>
		private void AddTableChildTours(Document doc, IEnumerable<Request> requests, IEnumerable<int> years)
		{
			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = _SIZE_22 });

			var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
			titleRequestRunPropertiesBold.AppendChild(new Bold());

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }));

			var table = new Table();
			table.AppendChild(tblProp.CloneNode(true));

			//первая строка (шапка)
			var row = new TableRow();
			var cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_1.ToString() }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
				new Run(titleRequestRunPropertiesBold.CloneNode(true),
					new Text("Год кампании"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_2.ToString() },
					new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) })),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
				new Run(titleRequestRunPropertiesBold.CloneNode(true),
					new Text("Вид услуги (путевка, сертификат, компенсация)"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth
					{
						Type = TableWidthUnitValues.Dxa,
						Width = (_COLUMN_3 + _COLUMN_4 + _COLUMN_5 /*+ SixthColumn*/).ToString()
					}),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new Run(titleRequestRunPropertiesBold.CloneNode(true),
					new Text("Организация отдыха и оздоровления (в случае предоставления путевки для отдыха и оздоровления), даты заезда"))));
			row.AppendChild(cell);
			table.AppendChild(row);

			var moneyTypes = new[]
			{
				(long?) TypeOfRestEnum.MoneyOn18, (long?) TypeOfRestEnum.MoneyOn3To7,
				(long?) TypeOfRestEnum.MoneyOn7To15, (long?) TypeOfRestEnum.MoneyOnInvalidOn4To17,
				(long?) TypeOfRestEnum.MoneyOnInvalidAndEscort4To17, (long?) TypeOfRestEnum.MoneyOnLimitationAndEscort4To17
			};

			//данные о предоставленных услугах

			foreach (var year in years)
			{

				Request request = requests.FirstOrDefault(req => req.YearOfRest.Year == year);
				if (!request.IsNullOrEmpty())
				{
					//string year = request?.YearOfRest?.Name;
					//if (request != null)
					//{
					//    if (request.ParentRequestId.HasValue && request.YearOfRest?.Year - 1 == request?.ParentRequest.YearOfRest?.Year)
					//    {
					//        year = $"{request.ParentRequest.YearOfRest?.Name} (дополнительная кампания)";
					//    }
					//}
					//else if (request == null)
					//{
					//    year = $"{request.ParentRequest.YearOfRest?.Name} (дополнительная кампания)";
					//}

					row = new TableRow();
					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_1.ToString() }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
						new Run(titleRequestRunProperties.CloneNode(true),
							new Text(year.ToString()))));
					row.AppendChild(cell);

					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _COLUMN_2.ToString() },
							new TableCellBorders(new BottomBorder
							{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
						new Run(titleRequestRunProperties.CloneNode(true),
							new Text(request?.TypeOfRest?.Name))));
					row.AppendChild(cell);

					cell = new TableCell();
					cell.Append(new TableCellProperties(
							new TableCellWidth
							{
								Type = TableWidthUnitValues.Dxa,
								Width = (_COLUMN_3 + _COLUMN_4 + _COLUMN_5 + _COLUMN_6).ToString()
							}),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell.AppendChild(new Paragraph(
						new Run(titleRequestRunProperties.CloneNode(true),
							new Text(request == null ? _SPACE :
								request.TourId.HasValue && !request.Tour.IsNullOrEmpty() && (request.TypeOfRest.Id != (long?)TypeOfRestEnum.Money && request.TypeOfRest.ParentId != (long?)TypeOfRestEnum.Money && request.TypeOfRest.Parent.ParentId != (long?)TypeOfRestEnum.Money)
								? $"{request.Tour.Hotels?.Name}, c {request.Tour.DateIncome.FormatEx(string.Empty)} по {request.Tour.DateOutcome.FormatEx(string.Empty)}"
								: !request.TourId.HasValue && (request.TypeOfRest.Id == (long?)TypeOfRestEnum.Money || request.TypeOfRest.ParentId == (long?)TypeOfRestEnum.Money || request.TypeOfRest.Parent.ParentId == (long?)TypeOfRestEnum.Money) && !request.Certificates.IsNullOrEmpty() && !request.Certificates.Where(c => c.Request.YearOfRestId == request.YearOfRestId).FirstOrDefault().IsNullOrEmpty()
								? $"{request.Certificates.Where(c => c.Request.YearOfRestId == request.YearOfRestId).FirstOrDefault().Place}, c {request.Certificates.Where(c => c.Request.YearOfRestId == request.YearOfRestId).FirstOrDefault().RestDateFrom.FormatEx(string.Empty)} по {request.Certificates.Where(c => c.Request.YearOfRestId == request.YearOfRestId).FirstOrDefault().RestDateTo.FormatEx(string.Empty)}"
								: (request.RequestOnMoney && !moneyTypes.Contains(request.TypeOfRestId)
									? "Осуществлен выбор сертификата на втором этапе заявочной кампании"
									: _SPACE)))));
					row.AppendChild(cell);
					table.AppendChild(row);
				}
			}

			doc.AppendChild(table);
		}

		private void DocumentHeaderRegistration(Document doc)
		{
			var titleProp = new RunProperties().SetFont().SetFontSize().Bold();
			doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center },
					new SpacingBetweenLines { After = _SPACING_BETWEEN_LINES_AFTER, Line = _SPACING_BETWEEN_LINES_LINE }),
				new Run(titleProp.CloneNode(true), new Text("ДЕПАРТАМЕНТ КУЛЬТУРЫ ГОРОДА МОСКВЫ"))));
			doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center },
					new SpacingBetweenLines { After = _SPACING_BETWEEN_LINES_AFTER, Line = _SPACING_BETWEEN_LINES_LINE }),
				new Run(titleProp.CloneNode(true),
					new Text("Государственное автономное учреждение культуры города Москвы"))));
			doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center },
					new SpacingBetweenLines { After = _SPACING_BETWEEN_LINES_AFTER, Line = _SPACING_BETWEEN_LINES_LINE }),
				new Run(titleProp.CloneNode(true), new Text("\"Московское агентство организации отдыха и туризма\""))));
			var pp = new ParagraphProperties(new Justification { Val = JustificationValues.Center },
				new SpacingBetweenLines { After = _SPACING_BETWEEN_LINES_AFTER, Line = _SPACING_BETWEEN_LINES_LINE });
			pp.AppendChild(new ParagraphBorders().BottomBorder("000000", 3));
			doc.AppendChild(new Paragraph(pp, new Run(titleProp.CloneNode(true), new Text("(ГАУК \"МОСГОРТУР\")"), new Break())));
		}

		/// <summary>
		/// Формирование таблицы перечня документов
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="Docs"></param>
		private void AddTableDocsList(Document doc, IEnumerable<string> Docs)
		{
			const string _rowonewight = "1000";
			const string _rowtwowight = "9266";

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }));
			var captionRunProperties = new RunProperties
			{
				RunFonts = new RunFonts
				{
					Ascii = "Times New Roman",
					HighAnsi = "Times New Roman",
					ComplexScript = "Times New Roman"
				},
				FontSize = new FontSize { Val = "20" },
				Italic = new Italic()
			};

			var table = new Table();
			table.AppendChild(tblProp.CloneNode(true));

			var _pgf = new List<string>();

			foreach (var _doc in Docs)
			{
				if (_doc.StartsWith("#"))
				{
					_pgf.Add(_doc.Substring(1));

					continue;
				}

				if (_doc.StartsWith("$"))
				{
					_pgf.Add(_doc.Substring(1));
				}
				else
				{
					if (_pgf.Any())
					{
						var _row = new TableRow();
						var _cell_one = new TableCell();
						_cell_one.Append(
							new TableCellProperties(
								new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _rowonewight }),
							new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
						_cell_one.AppendChild(new Paragraph(
							new ParagraphProperties(
								new Justification { Val = JustificationValues.Left }),
							new Run(captionRunProperties.CloneNode(true),
								new Text("  "))));
						_row.AppendChild(_cell_one);

						var _cell_two = new TableCell();
						_cell_two.Append(
							new TableCellProperties(
								new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _rowtwowight },
								new VerticalMerge { Val = MergedCellValues.Restart }
							),
							new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
						foreach (var _p in _pgf)
						{
							_cell_two.AppendChild(new Paragraph(
								new ParagraphProperties(
									new Justification { Val = JustificationValues.Left }),
								new Run(captionRunProperties.CloneNode(true),
									new Text(_p))));
						}

						_row.AppendChild(_cell_two);
						table.AppendChild(_row);

						_pgf.Clear();
					}

					var row = new TableRow();
					var cell_one = new TableCell();
					cell_one.Append(
						new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _rowonewight }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell_one.AppendChild(new Paragraph(
						new ParagraphProperties(
							new Justification { Val = JustificationValues.Left }),
						new Run(captionRunProperties.CloneNode(true),
							new Text("  "))));
					row.AppendChild(cell_one);

					var cell_two = new TableCell();
					cell_two.Append(
						new TableCellProperties(
							new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = _rowtwowight }),
						new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
					cell_two.AppendChild(new Paragraph(
						new ParagraphProperties(
							new Justification { Val = JustificationValues.Left }),
						new Run(captionRunProperties.CloneNode(true),
							new Text(_doc))));
					row.AppendChild(cell_two);
					table.AppendChild(row);
				}
			}

			doc.AppendChild(table);
		}

		private void SignBlockNotification(Document doc, Account account, string applicantName,
			bool NotificationAccepted = true)
		{
			//заглушка
			account = account ?? new Account();
			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = _SIZE_24 });

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));
			var captionRunProperties = new RunProperties().SetFont().SetFontSize("16");

			var table = new Table();
			table.AppendChild(tblProp.CloneNode(true));


			//отступ - нулевая строка
			{
				doc.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Both },
				new SpacingBetweenLines { After = _SIZE_20 })));
			}

			var row = new TableRow();

			var cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text("Уведомление выдал:"))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.Append(new TableCellProperties(
				new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" },
				new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom }
			));
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(account?.Name?.FormatEx()))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" },
					new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) })),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
				new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1804" },
				new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom }));
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(DateTime.Now.Date.FormatEx()))));
			row.AppendChild(cell);

			table.AppendChild(row);
			// -----------------------------------------------------------
			row = new TableRow();

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(_SPACE))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(captionRunProperties.CloneNode(true), new Text("(Ф.И.О. работника)"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true), captionRunProperties.CloneNode(true),
					new Text("(подпись)"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1804" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true), captionRunProperties.CloneNode(true),
					new Text("(дата)"))));
			row.AppendChild(cell);

			table.AppendChild(row);

			doc.AppendChild(table);

			if (NotificationAccepted)
			{
				table = new Table();
				//---------------------------------------------------------------------------
				row = new TableRow();

				table.AppendChild(tblProp.CloneNode(true));

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text("Уведомление принял:"))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(string.IsNullOrWhiteSpace(applicantName) ? "  " : applicantName))));
				row.AppendChild(cell);


				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1804" },
						new TableCellBorders(new BottomBorder
						{ Val = new EnumValue<BorderValues>(BorderValues.Single) })),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(DateTime.Now.Date.FormatEx()))));
				row.AppendChild(cell);

				table.AppendChild(row);

				// --------------------------------------------------------------------------------------
				row = new TableRow();

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(captionRunProperties.CloneNode(true), new Text("(Ф.И.О. заявителя)"))));
				row.AppendChild(cell);


				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(captionRunProperties.CloneNode(true), new Text("(подпись)"))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
				row.AppendChild(cell);

				cell = new TableCell();
				cell.Append(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1804" }),
					new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
				cell.AppendChild(new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
					new Run(captionRunProperties.CloneNode(true), new Text("(дата)"))));
				row.AppendChild(cell);

				table.AppendChild(row);

				doc.AppendChild(table);
			}
		}

		/// <summary>
		///  Блок подписией для ЛОК 2019
		/// </summary>
		private void SignBlockNotification2019(Document doc, Account account, string signText = null)
		{
			if (string.IsNullOrWhiteSpace(signText))
			{
				signText =
					"Подпись заявителя, подтверждающая получение уведомления о регистрации заявления на выплату компенсации за самостоятельно приобретенную родителями или иными законными представителями путевку для отдыха и оздоровления:";
			}

			SignWorkerBlock(doc, account);

			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = "22" });

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));
			var signRunProperties = new RunProperties().SetFont().SetFontSizeSupperscript();

			doc.AppendChild(
				new Paragraph(
					new ParagraphProperties(new Justification { Val = JustificationValues.Both },
						new SpacingBetweenLines { After = _SIZE_20 },
						new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
					new Run(titleRequestRunProperties.CloneNode(true),
						new Text(signText) { Space = SpaceProcessingModeValues.Preserve })
				));

			var table = new Table();
			//---------------------------------------------------------------------------
			var row = new TableRow();

			table.AppendChild(tblProp.CloneNode(true));

			var cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1731" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" },
					new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) })),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(_SPACE))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3931" },
					new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) })),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1804" },
					new TableCellBorders(new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single) })),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(DateTime.Now.Date.FormatEx()))));
			row.AppendChild(cell);

			table.AppendChild(row);

			// --------------------------------------------------------------------------------------
			row = new TableRow();

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(titleRequestRunProperties.CloneNode(true),
					new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2731" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(signRunProperties.CloneNode(true), new Text("(подпись)"))));
			row.AppendChild(cell);


			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2931" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(signRunProperties.CloneNode(true), new Text("(Ф.И.О. заявителя, расшифровка подписи)"))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "55" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(new Run(new Text(_SPACE))));
			row.AppendChild(cell);

			cell = new TableCell();
			cell.Append(new TableCellProperties(
					new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1804" }),
				new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Bottom });
			cell.AppendChild(new Paragraph(
				new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(signRunProperties.CloneNode(true), new Text("(дата)"))));
			row.AppendChild(cell);

			table.AppendChild(row);

			doc.AppendChild(table);
		}

		/// <summary>
		///     Уведомление о регистрации заявления о предоставлении услуг отдыха и оздоровления при альтернативной кампании
		/// </summary>
		private IDocument AlternateCompanyNotification(Request request, Account account, bool forMpguPortal = false, bool isDecline = false, string time = "", List<Request> requests = null)
		{
			var youth = request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps;

			var isCert = request.DeclineReasonId == (long)DeclineReasonEnum.CertificateIssued ||
							request.TypeOfRestId == (long)TypeOfRestEnum.Money ||
							request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn18 ||
							request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn3To7 ||
							request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOn7To15 ||
							request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidOn4To17 ||
							request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnInvalidAndEscort4To17 ||
							request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnLimitationAndEscort4To17 ||
								 request.TypeOfRestId == (long)TypeOfRestEnum.MoneyOnBenefits
								 ||
								 request.TypeOfRest.ParentId == (long)TypeOfRestEnum.Money; ;

			var isNewChoise = request.StatusId == (long)StatusEnum.DecisionMakingCovid;

			//if (forMpguPortal)
			//{
			//    return PDFDocuments.PdfProcessor.NotificationBasicRegistration(request, youth, account);
			//}

			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					if (isDecline)
					{
						var elems = new List<OpenXmlElement>();
						elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold());
						elems.Add(new Text("Уведомление об отказе от всех вариантов организации отдыха"));
						elems.Add(new Break());
						elems.Add(new Text("и оздоровления"));

						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(elems)
						));
					}
					else
					if (isNewChoise)
					{
						var elems = new List<OpenXmlElement>();
						elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold());
						elems.Add(new Text("Уведомление о необходимости выбора конкретной организации отдыха"));
						elems.Add(new Break());
						elems.Add(new Text("и оздоровления"));

						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(elems)
						));
					}
					else
					if (!isCert)
					{
						var elems = new List<OpenXmlElement>();
						elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold());
						elems.Add(new Text("Уведомление об осуществлении выбора варианта организации отдыха"));
						elems.Add(new Break());
						elems.Add(new Text("и оздоровления (бесплатная путевка для отдыха и оздоровления)"));

						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(elems)
						));
					}
					else
					{
						var elems = new List<OpenXmlElement>();
						elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold());
						elems.Add(new Text("Уведомление об осуществлении выбора варианта организации отдыха"));
						elems.Add(new Break());
						elems.Add(new Text("и оздоровления (сертификат на отдых и оздоровление)"));
						doc.AppendChild(new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
							new Run(elems)
						));
					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var documentRunPropertiesItalic = new RunProperties().SetFont().SetFontSize(_SIZE_16);
					documentRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));

					if (isCert)
						doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text("сертификат на отдых и оздоровление"))));
					else
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Both },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Вид отдыха: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(youth
										? "бесплатная путевка для отдыха и оздоровления"
										: request.TypeOfRest?.Name))));

					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Данные ребёнка: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text($"{child.BenefitType?.Name}"))));

					}

					if (!isDecline)
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения заявления: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text("Услуга оказана."))));

						if (isCert)
						{
							string certificateNumbers = "";
							foreach (var r in requests)
								certificateNumbers += (r.CertificateNumber) + ", ";
							doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Номер сертификата: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(certificateNumbers))));
						}
						else
						{
							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Дата и время выбора варианта организации отдыха и оздоровления: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{request.CertificateDate}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Номер путевки: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{request.CertificateNumber}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Организация отдыха и оздоровления: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{request.PlaceOfRest.Name}"))));

							doc.AppendChild(
								new Paragraph(
									new ParagraphProperties(new Justification { Val = JustificationValues.Left },
										new SpacingBetweenLines { After = _SIZE_20 }),
									new Run(titleRequestRunPropertiesBold.CloneNode(true),
										new Text("Время отдыха: ")
										{ Space = SpaceProcessingModeValues.Preserve }),
									new Run(titleRequestRunPropertiesItalic.CloneNode(true),
										new Text($"{request.TimeOfRest.Name}"))
								));
						}
					}
					else
					{
						//var time1 = time.TryParseDateDdMmYyyy();
						doc.AppendChild(
							   new Paragraph(
								   new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									   new SpacingBetweenLines { After = _SIZE_20 }),
								   new Run(titleRequestRunPropertiesBold.CloneNode(true),
									   new Text("Дата и время принятия решения об отказе от всех вариантов организации отдыха и оздоровления: ")
									   { Space = SpaceProcessingModeValues.Preserve }),
								   new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									   new Text(time))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Результат рассмотрения заявления: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text("отказ в предоставлении услуг отдыха и оздоровления"),
									new Break())));


						var elems = new List<OpenXmlElement>();
						elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_22).Bold());
						elems.Add(new Text("Я подтверждаю отказ от всех вариантов организации отдыха и оздоровления."));
						elems.Add(new Break());
						elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_22).Bold());
						elems.Add(new Text("Поставлен(а) в известность, что такой отказ не может являться основанием для отказа в предоставлении услуг отдыха и оздоровления в последующие годы."));

						doc.AppendChild(new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new Indentation { FirstLine = _FIRST_LINE_INDENTATION_600.ToString() }),
								new Run(elems)
							));


					}

					SignBlockNotification2020(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о решении альтернативной кампании" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		///     Уведомление о перевыборе времени и места отдыха в заявлении о предоставлении услуг отдыха и оздоровления при альтернативной кампании
		/// </summary>
		private IDocument AlternateCompanyRechoiseNotification(Request request, Account account, bool forMpguPortal = false, string time = "")
		{
			var youth = request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestCamps ||
						request.TypeOfRestId == (long)TypeOfRestEnum.YouthRestOrphanCamps;

			//if (forMpguPortal)
			//{
			//    return PDFDocuments.PdfProcessor.NotificationBasicRegistration(request, youth, account);
			//}

			using (var ms = new MemoryStream())
			{
				using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
				{
					var mainPart = wordDocument.AddMainDocumentPart();
					var doc = new Document(new Body());

					DocumentHeaderRegistration(doc);

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));


					var elems = new List<OpenXmlElement>();
					elems.Add(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold());
					elems.Add(new Text("Уведомление о необходимости выбора конкретной организации отдыха"));
					elems.Add(new Break());
					elems.Add(new Text("и оздоровления"));

					doc.AppendChild(new Paragraph(
						new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
						new Run(elems)
					));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(new RunProperties().SetFont().SetFontSize(_SIZE_28).Bold(), new Text(_SPACE))));

					var titleRequestRunProperties = new RunProperties().SetFont().SetFontSize();

					var titleRequestRunPropertiesBold = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesBold.AppendChild(new Bold());

					var titleRequestRunPropertiesItalic = titleRequestRunProperties.CloneNode(true);
					titleRequestRunPropertiesItalic.AppendChild(new Italic());

					var documentRunPropertiesItalic = new RunProperties().SetFont().SetFontSize(_SIZE_16);
					documentRunPropertiesItalic.AppendChild(new Italic());

					var applicant = request.Applicant ??
									new Applicant { DocumentType = new DocumentType { Name = string.Empty } };

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время регистрации заявления: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.DateRequest.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Номер заявления: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.RequestNumber.FormatEx()))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Данные заявителя: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(
									$"{applicant.LastName} {applicant.FirstName} {applicant.MiddleName}, {applicant.DateOfBirth.FormatExGR(string.Empty)}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Контактная информация: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(applicant.Phone + ", " + applicant.Email))));


					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Both },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Вид услуги: ") { Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TypeOfRest?.Name))));

					foreach (var child in request.Child?.Where(c => !c.IsDeleted) ?? new List<Child>())
					{
						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Данные ребёнка: ") { Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text(
										$"{child.LastName} {child.FirstName} {child.MiddleName}, {child.DateOfBirth.FormatExGR(string.Empty)}"))));

						doc.AppendChild(
							new Paragraph(
								new ParagraphProperties(new Justification { Val = JustificationValues.Left },
									new SpacingBetweenLines { After = _SIZE_20 }),
								new Run(titleRequestRunPropertiesBold.CloneNode(true),
									new Text("Льготная категория: ")
									{ Space = SpaceProcessingModeValues.Preserve }),
								new Run(titleRequestRunPropertiesItalic.CloneNode(true),
									new Text($"{child.BenefitType?.Name}"))));

					}

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дата и время принятия решения о выборе варианта предоставления бесплатной путевки для выездного отдыха: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(time))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Приоритетное время отдыха: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.TimeOfRest.Name))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дополнительное время отдыха: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text($"{request.TimesOfRest.FirstOrDefault().TimeOfRest.Name}, {request.TimesOfRest.LastOrDefault().TimeOfRest.Name}"))));

					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Приоритетное направление отдыха: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text(request.PlaceOfRest.Name))));
					doc.AppendChild(
						new Paragraph(
							new ParagraphProperties(new Justification { Val = JustificationValues.Left },
								new SpacingBetweenLines { After = _SIZE_20 }),
							new Run(titleRequestRunPropertiesBold.CloneNode(true),
								new Text("Дополнительное направление отдыха: ")
								{ Space = SpaceProcessingModeValues.Preserve }),
							new Run(titleRequestRunPropertiesItalic.CloneNode(true),
								new Text($"{request.PlacesOfRest.FirstOrDefault().PlaceOfRest.Name}, {request.PlacesOfRest.LastOrDefault().PlaceOfRest.Name}"))));


					SignBlockNotification2020(doc, account, "Исполнитель:");

					mainPart.Document = doc;
				}

				return new DocumentResult
				{
					FileBody = ms.ToArray(),
					FileName = "Уведомление о решении альтернативной кампании" + Extension,
					MimeType = MimeType,
					MimeTypeShort = Extension
				};
			}
		}

		/// <summary>
		///     добавление заголовка
		/// </summary>
		private void AppendHeader(Document doc)
		{
			var titleParagraphProperties = new ParagraphProperties();
			var blockTitleRunProperties = new RunProperties();
			blockTitleRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			blockTitleRunProperties.AppendChild(new FontSize { Val = "24" });
			titleParagraphProperties.AppendChild(new Justification { Val = JustificationValues.Center });
			titleParagraphProperties.AppendChild(new Indentation { Left = "6000" });

			doc.AppendChild(new Paragraph(titleParagraphProperties,
				new Run(blockTitleRunProperties,
					new Text(
						"Государственному автономному учреждению культуры города Москвы «Московское агентство организации отдыха и туризма»"))));

			doc.AppendChild(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
				new Run(MakeTitleRequestRunProperties(),
					new Text(
						"Заявление о предоставлении услуг отдыха и оздоровления (в части сведений, заполняемых в период проведения первого этапа заявочной кампании в целях организации отдыха и оздоровления)"))));

		}

		private RunProperties MakeTitleRequestRunProperties()
		{
			var titleRequestRunProperties = new RunProperties();
			titleRequestRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			titleRequestRunProperties.AppendChild(new FontSize { Val = "28" });
			titleRequestRunProperties.AppendChild(new Bold());
			return titleRequestRunProperties;
		}

		/// <summary>
		///     добавление блока
		/// </summary>
		private void AppendBlock(Document doc, string title, IEnumerable<Tuple<string, string>> items)
		{
			var titleParagraphProperties = new ParagraphProperties();
			var blockTitleRunProperties = new RunProperties();
			blockTitleRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			blockTitleRunProperties.AppendChild(new FontSize { Val = "28" });
			blockTitleRunProperties.AppendChild(new Bold());
			titleParagraphProperties.AppendChild(new SpacingBetweenLines
			{
				Before = "300",
				After = "100",
				LineRule = LineSpacingRuleValues.Exact
			});
			titleParagraphProperties.AppendChild(new KeepNext());

			doc.AppendChild(new Paragraph(titleParagraphProperties, new Run(blockTitleRunProperties, new Text(title))));

			var table = new Table();

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));

			table.AppendChild(tblProp);

			if (items != null)
			{
				foreach (var item in items)
				{
					var recordTitle = item.Item1;
					var recordVal = item.Item2;
					var row = new TableRow();
					var cell1 = new TableCell();
					var cell2 = new TableCell();
					var run1Properties = new RunProperties();
					var run2Properties = new RunProperties();

					run1Properties.AppendChild(new RunFonts
					{
						Ascii = "Times New Roman",
						HighAnsi = "Times New Roman",
						ComplexScript = "Times New Roman"
					});
					run1Properties.AppendChild(new FontSize { Val = "24" });

					run2Properties.AppendChild(new RunFonts
					{
						Ascii = "Times New Roman",
						HighAnsi = "Times New Roman",
						ComplexScript = "Times New Roman"
					});
					run2Properties.AppendChild(new FontSize { Val = "24" });
					run2Properties.AppendChild(new Bold());

					cell1.AppendChild(new TableCellProperties(
						new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3500" }));
					cell1.AppendChild(new Paragraph(new Run(run1Properties, new Text(recordTitle))));
					cell2.AppendChild(new Paragraph(new ParagraphProperties(new Indentation { Left = "500" }),
						new Run(run2Properties, new Text(recordVal))));

					row.AppendChild(cell1);
					row.AppendChild(cell2);
					table.AppendChild(row);
				}

				doc.AppendChild(table);
			}
		}

		/// <summary>
		///     добавление заглавия
		/// </summary>
		private void AppendTitle(Document doc, string title)
		{
			var titleParagraphProperties = new ParagraphProperties();
			var blockTitleRunProperties = new RunProperties();
			blockTitleRunProperties.AppendChild(new RunFonts
			{
				Ascii = "Times New Roman",
				HighAnsi = "Times New Roman",
				ComplexScript = "Times New Roman"
			});
			blockTitleRunProperties.AppendChild(new FontSize { Val = "28" });
			blockTitleRunProperties.AppendChild(new Bold());
			titleParagraphProperties.AppendChild(new SpacingBetweenLines
			{
				Before = "300",
				LineRule = LineSpacingRuleValues.Exact
			});
			titleParagraphProperties.AppendChild(new KeepNext());

			doc.AppendChild(new Paragraph(titleParagraphProperties, new Run(blockTitleRunProperties, new Text(title))));

			var table = new Table();

			var tblProp = new TableProperties(
				new TableBorders(
					new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) },
					new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.None) }));

			table.AppendChild(tblProp);
		}

		private void AppendFooterNotificationOfRegistration(Document doc)
		{
			var assembly = Assembly.Load("RestChild.Templates");
			var resourceName = "RestChild.Templates.NotificationRegistrationFooter" + Extension;
			var chunk = doc.MainDocumentPart.AddAlternativeFormatImportPart(
				AlternativeFormatImportPartType.WordprocessingML, "AltChunkIdFooter");
			using (var stream = assembly.GetManifestResourceStream(resourceName))
			{
				chunk.FeedData(stream);
			}

			var altChunk = new AltChunk { Id = "AltChunkIdFooter" };
			doc.InsertAfter(altChunk, doc.Elements().Last());
		}

		/// <summary>
		///     Наименование заявителя
		/// </summary>
		private string GetStatusApplicantName(string statusApplicant)
		{
			switch (statusApplicant)
			{
				case "1":
					statusApplicant = "Отец";
					break;
				case "2":
					statusApplicant = "Мать";
					break;
				case "3":
					statusApplicant = "Законный представитель";
					break;
				case "4":
					statusApplicant = "Доверенное лицо";
					break;
				case "5":
					statusApplicant = "Лицо из числа детей-сирот и детей, оставшихся без попечения родителей";
					break;
			}

			return statusApplicant;
		}
		#endregion
	}
}


//Свитчер документов, в результате он похудел колоссально
namespace RestChild.DocumentGeneration
{
	/// <summary>
	///     маршрутизация генерации документа
	/// </summary>
	public static class DocumentSwitch
	{
		/// <summary>
		///     генерация документа
		/// </summary>
		public static ICshedDocument Switch(Account account, Request request, DocumentGenerationEnum.DocType docType,
			string documentPath, bool requestOnMoney = false, long? sendStatusId = null)
		{
			IBaseDocumentProcessor docProc = null;

			switch (docType)
			{
				case DocumentGenerationEnum.DocType.Doc:
					docProc = new DocProc();
					break;
				case DocumentGenerationEnum.DocType.Pdf:
					docProc = new PdfProc();
					break;
				default:
					return null;
			}

			ICshedDocument doc;

			try
			{
				if (documentPath.Equals(DocumentGenerationEnum.SaveCertificateToRequest))
				{
					doc = (ICshedDocument)docProc.SaveCertificateToRequest(request, sendStatusId);

					if (doc != null)
					{
						if (requestOnMoney)
						{
							doc.RequestFileTypeId = (long)RequestFileTypeEnum.CertificateOnPayment;
							doc.FileName = "Сертификат.pdf";
						}
						else
						{
							doc.RequestFileTypeId = (long)RequestFileTypeEnum.CertificateOnRest;
							doc.FileName = "Путёвка.pdf";
						}
					}
				}
				else if (documentPath.Equals(DocumentGenerationEnum.NotificationDeclineRegistryReg))
				{
					doc = (ICshedDocument)docProc.NotificationDeclineRegistryReg(request, account);
				}
				else if (documentPath.Equals(DocumentGenerationEnum.NotificationRegistration))
				{
					doc = (ICshedDocument)docProc.NotificationRegistration(request, account);
				}
				else if (documentPath.Equals(DocumentGenerationEnum.NotificationRefuse))
				{
					doc = (ICshedDocument)docProc.NotificationRefuse(request, account);
				}
				else if (documentPath.Equals(DocumentGenerationEnum.NotificationWaitApplicant))
				{
					doc = (ICshedDocument)docProc.NotificationWaitApplicant(request, account);
				}
				else if (documentPath.Equals(DocumentGenerationEnum.NotificationAboutDecision))
				{
					doc = (ICshedDocument)docProc.NotificationAboutDecision(request, account);
				}
				else if (documentPath.Equals(DocumentGenerationEnum.NotificationOfNeedToChoose))
				{
					// Уведомление о необходимости выбора организации отдыха и оздоровления
					doc = (ICshedDocument)docProc.NotificationOrgChoose(request, account);
				}
				else if (documentPath.Equals(DocumentGenerationEnum.NotificationAttendantChange))
				{
					// Уведомление о смене сопровождающего
					doc = (ICshedDocument)docProc.NotificationAttendantChange(request, account);
				}
				else if (documentPath.Equals(DocumentGenerationEnum.NotificationDeclineAccepted))
				{
					// Уведомление об отказе от путёвки/сертификата
					doc = (ICshedDocument)docProc.NotificationDeclineAccepted(request, account);
				}
				else if (documentPath.Equals(DocumentGenerationEnum.NotificationDeclineNotAccepted))
				{
					// Уведомление об отказе в отказе от путёвки/сертификата
					doc = (ICshedDocument)docProc.NotificationDeclineNotAccepted(request, account);
				}
				else
				{
					return null;
				}

				//загадочное поведение генерации файла
				if (doc?.FileBody == null)
				{
					return new CshedDocumentResult();
				}

				if (!doc.RequestFileTypeId.HasValue || doc.RequestFileTypeId.Value < 1)
				{
					doc.RequestFileTypeId = (long)RequestFileTypeEnum.Notifications;
				}

				return doc;
			}
			catch (NotImplementedException ex) //не во всех наследниках реализованы методы генерирующие документы
			{
				return null;
			}
		}
	}
}
