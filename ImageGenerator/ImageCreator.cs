/*
Здесь полностью мной написан генератор изображение. Смысл был таков, что есть место отдыха, и на его картинку должно наложиться было
значок глухой, слепой, хромой, полный инвалид. Причем вариант с передачей параметров на фронт отпадал, ибо тогда требовалось
дорабатывать Портал со стороны МГТ, на что заказчик нам сказал, что так не пойдет ищите вариант без доработок на их стороне.
Поэтому картинка генерируется и сразу отдается готовая
 */

//рисовальщик
namespace RestChild.Web.Common
{
	public class AccessableEnvImageCreator
	{
		/// <summary>
		/// Ширина по умолчанию к которой надо привести оригинальное изображение
		/// </summary>
		private const int _IMAGE_DEFAULT_WIDTH_TO = 700; //px

		/// <summary>
		/// Ширина по умолчанию к которой надо привести иконку
		/// </summary>
		private const int _ICON_DEFAULT_WIDTH_TO_ICON = 100; //px

		/// <summary>
		/// Коренная папка иконок доступности среды
		/// </summary>
		private const string _ICONS_ROOT_FOLDER = "Content/images/RestricionGroups";

		/// <summary>
		/// Подпапка частично-доступной среды 
		/// </summary>
		private const string _ICONS_PARTLY_ACCESSABLE_ENV_SUB_FOLDER = "PartlyAccessableEnvironment";

		//именование файлов иконок
		private const string _ICON_DEAF = "Deaf.png",
							 _ICON_LAME = "Lame.png",
							 _ICON_BLIND = "Blind.png",
							 _ICON_ACCESSABLE_ENV = "AccessableEnvironment.png";

		/// <summary>
		/// Размер по умолчанию отступа от границ иконок
		/// </summary>
		private const int _ICON_MARGIN = 5; //px

		/// <summary>
		/// Размер отступа от границ иконок
		/// </summary>
		public int IconMargin { get; set; } = _ICON_MARGIN;

		/// <summary>
		/// Ширина к которой надо привести оригинальное изображение
		/// </summary>
		public int ImageWidthTo { get; set; } = _IMAGE_DEFAULT_WIDTH_TO;

		/// <summary>
		/// Ширина к которой надо привести иконки
		/// </summary>
		public int IconWidthTo { get; set; } = _ICON_DEFAULT_WIDTH_TO_ICON;

		/// <summary>
		/// Глухой
		/// </summary>
		public bool Lame { get; set; } = false;

		/// <summary>
		/// Слепой
		/// </summary>
		public bool Blind { get; set; } = false;

		/// <summary>
		/// Хромой
		/// </summary>
		public bool Deaf { get; set; } = false;

		public bool AccessableEnv { get; set; } = false;

		/// <summary>
		/// Путь к оригинально
		/// </summary>
		private string _origImgPath;

		/// <summary>
		/// Сторона
		/// </summary>
		private enum Sides
		{
			Height,
			Width,
		}

		public AccessableEnvImageCreator(string imgPath, bool deaf = false, bool lame = false, bool blind = false, bool accessableEnv = false)
		{
			_origImgPath = imgPath;
			Lame = lame;
			Deaf = deaf;
			Blind = blind;
			AccessableEnv = accessableEnv;
		}

		public byte[] GetImage()
		{
			Bitmap bmp = ReduceBitmapSizeProportionately(new Bitmap(_origImgPath), ImageWidthTo, Sides.Width);
			Bitmap deafBmp = ReduceBitmapSizeProportionately(new Bitmap(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _ICONS_ROOT_FOLDER, _ICONS_PARTLY_ACCESSABLE_ENV_SUB_FOLDER, _ICON_DEAF)), IconWidthTo, Sides.Width);
			Bitmap lameBmp = ReduceBitmapSizeProportionately(new Bitmap(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _ICONS_ROOT_FOLDER, _ICONS_PARTLY_ACCESSABLE_ENV_SUB_FOLDER, _ICON_LAME)), IconWidthTo, Sides.Width);
			Bitmap blindBmp = ReduceBitmapSizeProportionately(new Bitmap(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _ICONS_ROOT_FOLDER, _ICONS_PARTLY_ACCESSABLE_ENV_SUB_FOLDER, _ICON_BLIND)), IconWidthTo, Sides.Width);
			Bitmap accessableEnvBmp = ReduceBitmapSizeProportionately(new Bitmap(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _ICONS_ROOT_FOLDER, _ICON_ACCESSABLE_ENV)), IconWidthTo, Sides.Width);
			List<Bitmap> iconsToDraw = new List<Bitmap>();
			if (Deaf) iconsToDraw.Add(deafBmp);
			if (Lame) iconsToDraw.Add(lameBmp);
			if (Blind) iconsToDraw.Add(blindBmp);
			if (AccessableEnv) iconsToDraw.Add(accessableEnvBmp);

			using (Graphics g = Graphics.FromImage(bmp))
			{
				Point startPoint = new Point(bmp.Size.Width, bmp.Size.Height);

				iconsToDraw.ForEach(icon =>
				{
					startPoint.X -= (icon.Width + IconMargin * 2);
					startPoint.Y = bmp.Size.Height - (icon.Height + IconMargin * 2);
					g.DrawImage(icon, startPoint);
				});
			}

			using (MemoryStream ms = new MemoryStream())
			{
				bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
				return ms.ToArray();
			}

		}

		private Bitmap ReduceBitmapSizeProportionately(Bitmap originalBmp, int sizeTo, Sides side)
		{
			Size origSize = originalBmp.Size;
			Size size = origSize;

			switch (side)
			{
				case Sides.Height:
					size = new Size((int)Math.Round(origSize.Width * (decimal)(sizeTo / origSize.Height), 0), sizeTo);
					break;
				case Sides.Width:
					size = new Size(sizeTo, (int)Math.Round(origSize.Height * ((decimal)sizeTo / (decimal)origSize.Width), 0));
					break;
			}

			return new Bitmap(originalBmp, size);
		}
	}
}

//Http обработчик на запрос о картинке, он уже был я только модефицировал
//до
namespace RestChild.Web.Handlers
{
	public class UploadHandlerBase : IHttpHandler
	{
		private readonly JavaScriptSerializer _js;

		public UploadHandlerBase()
		{
			_js = new JavaScriptSerializer { MaxJsonLength = 41943040 };
		}

		protected virtual string StorageRoot => Path.Combine(Settings.Default.StorageRootPath);

		public bool IsReusable => false;

		public void ProcessRequest(HttpContext context)
		{
			context.Response.AddHeader("Pragma", "no-cache");
			context.Response.AddHeader("Cache-Control", "private, no-cache");

			HandleMethod(context);
		}

		// Handle request based on method
		private void HandleMethod(HttpContext context)
		{
			//var accountId = RestChild.Web.Controllers.Security.GetCurrentAccountId();
			//if (!accountId.HasValue && context.Request["sk"] != Settings.Default.SecretKey)
			//{
			//	context.Response.ClearHeaders();
			//	context.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
			//	return;
			//}

			switch (context.Request.HttpMethod)
			{
				case "HEAD":
				case "GET":
					if (RemoteFilename(context))
					{
						DeliverChedFile(context);
					}
					else if (GivenFilename(context))
					{
						DeliverFile(context);
					}
					else
					{
						context.Response.ClearHeaders();
						context.Response.StatusCode = (int)HttpStatusCode.MethodNotAllowed;
					}

					break;
				case "POST":
				case "PUT":
					UploadFile(context);
					break;

				case "DELETE":
					DeleteFile(context);
					break;

				case "OPTIONS":
					ReturnOptions(context);
					break;

				default:
					context.Response.ClearHeaders();
					context.Response.StatusCode = (int)HttpStatusCode.MethodNotAllowed;
					break;
			}
		}

		private static void ReturnOptions(HttpContext context)
		{
			context.Response.AddHeader("Allow", "DELETE,GET,HEAD,POST,PUT,OPTIONS");
			context.Response.StatusCode = (int)HttpStatusCode.OK;
		}

		// Delete file from the server
		private void DeleteFile(HttpContext context)
		{
			var filePath = StorageRoot + context.Request["f"];
			if (File.Exists(filePath))
			{
				//File.Delete(filePath);
			}
		}

		// Upload file to the server
		private void UploadFile(HttpContext context)
		{
			var statuses = new List<FilesStatus>();
			var headers = context.Request.Headers;

			if (string.IsNullOrEmpty(headers["Content-Range"]))
			{
				UploadWholeFile(context, statuses);
			}
			else
			{
				UploadPartialFile(context, statuses);
			}

			WriteJsonIframeSafe(context, statuses);
		}

		// Upload partial file
		private void UploadPartialFile(HttpContext context, List<FilesStatus> statuses)
		{
			if (context.Request.Files.Count != 1)
				throw new HttpRequestValidationException(
					"Attempt to upload chunked file containing more than one fragment per request");
			var inputStream = context.Request.Files[0].InputStream;
			var fileName = Path.GetFileName(context.Request.Files[0].FileName);

			var headers = context.Request.Headers;
			var realName = headers["X-FileName"];
			realName = string.IsNullOrEmpty(realName)
				? Guid.NewGuid() + Path.GetExtension(context.Request.Files[0].FileName)
				: realName;

			var fullName = StorageRoot + Path.GetFileName(realName);
			var contentRange = headers["Content-Range"].Replace("bytes ", string.Empty);
			var size = long.Parse(contentRange.Substring(contentRange.IndexOf("/", StringComparison.CurrentCulture) + 1));
			var startposition = long.Parse(contentRange.Substring(0, contentRange.IndexOf("-", StringComparison.CurrentCulture)));

			using (var fs = new FileStream(fullName, FileMode.OpenOrCreate, FileAccess.Write))
			{
				fs.SetLength(size);
				fs.Seek(startposition, SeekOrigin.Begin);

				var buffer = new byte[1024];

				var l = inputStream.Read(buffer, 0, 1024);
				while (l > 0)
				{
					fs.Write(buffer, 0, l);
					l = inputStream.Read(buffer, 0, 1024);
				}
				fs.Flush();
				fs.Close();
			}
			statuses.Add(new FilesStatus(fileName, new FileInfo(fullName)));
		}

		// Upload entire file
		private void UploadWholeFile(HttpContext context, List<FilesStatus> statuses)
		{
			for (var i = 0; i < context.Request.Files.Count; i++)
			{
				var file = context.Request.Files[i];

				var realFileName = Guid.NewGuid() + Path.GetExtension(file.FileName);

				var fullPath = StorageRoot + realFileName;

				file.SaveAs(fullPath);

				var fullName = Path.GetFileName(file.FileName);
				statuses.Add(new FilesStatus(fullName, realFileName, file.ContentLength, fullPath));
			}
		}

		private void WriteJsonIframeSafe(HttpContext context, List<FilesStatus> statuses)
		{
			context.Response.AddHeader("Vary", "Accept");
			try
			{
				context.Response.ContentType = context.Request["HTTP_ACCEPT"].Contains("application/json") ? "application/json" : "text/plain";
			}
			catch
			{
				context.Response.ContentType = "text/plain";
			}

			var jsonObj = _js.Serialize(statuses.ToArray());
			context.Response.Write(jsonObj);
		}

		private static bool RemoteFilename(HttpContext context)
		{
			return !string.IsNullOrEmpty(context.Request["r"]);
		}

		private static bool GivenFilename(HttpContext context)
		{
			return !string.IsNullOrEmpty(context.Request["f"]);
		}

		public static void CopyStreamTo(Stream src, Stream dest)
		{
			int num = src.CanSeek ? Math.Min((int)(src.Length - src.Position), 8192) : 8192;
			byte[] array = new byte[num];
			int num2;
			do
			{
				num2 = src.Read(array, 0, array.Length);
				dest.Write(array, 0, num2);
			}
			while (num2 != 0);
		}

		private void DeliverAisoFile(HttpContext context)
		{
			var filename = context.Request["f"];
			var filetitle = context.Request["t"];
			filetitle = string.IsNullOrEmpty(filetitle) ? filename : filetitle;
			var req = WebRequest.Create(WebConfigurationManager.AppSettings["AisoUrl"] + "/Upload.ashx?r=1&f=" + filename);
			using (var resp = req.GetResponse())
			{
				using (var stream = resp.GetResponseStream())
				{
					var buffer = new byte[1000000];
					var bytesRead = 0;
					context.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + filetitle + "\"");
					context.Response.ContentType = "application/octet-stream";
					context.Response.ClearContent();
					while ((bytesRead = stream?.Read(buffer, 0, buffer.Length) ?? 0) != 0)
					{
						context.Response.OutputStream.Write(buffer, 0, bytesRead);
					}
				}
			}
		}

		private void DeliverChedFile(HttpContext context)
		{
			var filename = context.Request["f"];
			var filetitle = context.Request["t"];
			filetitle = string.IsNullOrEmpty(filetitle) ? filename : filetitle;

			var data = DeliverChedFileStatic(context, filename, filetitle);

			context.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + filetitle + "\"");
			context.Response.ContentType = "application/octet-stream";
			context.Response.ClearContent();
			context.Response.BinaryWrite(data);
		}

		internal static byte[] DeliverChedFileStatic(HttpContext context, string filename, string filetitle)
		{
			using (var csClient = new CustomWebServiceImplClient())
			{
				if (csClient.ClientCredentials != null)
				{
					csClient.ClientCredentials.UserName.UserName = ConfigurationManager.AppSettings["CshedLogin"];
					csClient.ClientCredentials.UserName.Password = ConfigurationManager.AppSettings["CshedPass"];
				}

				return csClient.GetDocumentData(new GetDocumentRequest { DocumentId = filename });
			}
		}

		private void DeliverFile(HttpContext context)
		{
			var filename = context.Request["f"];
			var filetitle = context.Request["t"];
			filetitle = string.IsNullOrEmpty(filetitle) ? filename : filetitle;
			var filePath = StorageRoot + filename;

			if (File.Exists(filePath))
			{
				context.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + filetitle + "\"");
				context.Response.ContentType = "application/octet-stream";
				context.Response.ClearContent();
				context.Response.TransmitFile(filePath);
			}
			else
				context.Response.StatusCode = (int)HttpStatusCode.NotFound;
		}
	}
}

//после
namespace RestChild.Web.Handlers
{
	public class UploadHandlerBase : IHttpHandler
	{
		private readonly JavaScriptSerializer _js;

		public UploadHandlerBase()
		{
			_js = new JavaScriptSerializer { MaxJsonLength = 41943040 };
		}

		protected virtual string StorageRoot => Path.Combine(Settings.Default.StorageRootPath);

		public bool IsReusable => false;

		public void ProcessRequest(HttpContext context)
		{
			context.Response.AddHeader("Pragma", "no-cache");
			context.Response.AddHeader("Cache-Control", "private, no-cache");

			HandleMethod(context);
		}

		// Handle request based on method
		private void HandleMethod(HttpContext context)
		{
			//var accountId = RestChild.Web.Controllers.Security.GetCurrentAccountId();
			//if (!accountId.HasValue && context.Request["sk"] != Settings.Default.SecretKey)
			//{
			//	context.Response.ClearHeaders();
			//	context.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
			//	return;
			//}

			switch (context.Request.HttpMethod)
			{
				case "HEAD":
				case "GET":
					if (RemoteFilename(context))
					{
						DeliverChedFile(context);
					}
					else if (GivenFilename(context))
					{
						DeliverFile(context);
					}
					else
					{
						context.Response.ClearHeaders();
						context.Response.StatusCode = (int)HttpStatusCode.MethodNotAllowed;
					}

					break;
				case "POST":
				case "PUT":
					UploadFile(context);
					break;

				case "DELETE":
					DeleteFile(context);
					break;

				case "OPTIONS":
					ReturnOptions(context);
					break;

				default:
					context.Response.ClearHeaders();
					context.Response.StatusCode = (int)HttpStatusCode.MethodNotAllowed;
					break;
			}
		}

		private static void ReturnOptions(HttpContext context)
		{
			context.Response.AddHeader("Allow", "DELETE,GET,HEAD,POST,PUT,OPTIONS");
			context.Response.StatusCode = (int)HttpStatusCode.OK;
		}

		// Delete file from the server
		private void DeleteFile(HttpContext context)
		{
			var filePath = StorageRoot + context.Request["f"];
			if (File.Exists(filePath))
			{
				//File.Delete(filePath);
			}
		}

		// Upload file to the server
		private void UploadFile(HttpContext context)
		{
			var statuses = new List<FilesStatus>();
			var headers = context.Request.Headers;

			if (string.IsNullOrEmpty(headers["Content-Range"]))
			{
				UploadWholeFile(context, statuses);
			}
			else
			{
				UploadPartialFile(context, statuses);
			}

			WriteJsonIframeSafe(context, statuses);
		}

		// Upload partial file
		private void UploadPartialFile(HttpContext context, List<FilesStatus> statuses)
		{
			if (context.Request.Files.Count != 1)
				throw new HttpRequestValidationException(
					"Attempt to upload chunked file containing more than one fragment per request");
			var inputStream = context.Request.Files[0].InputStream;
			var fileName = Path.GetFileName(context.Request.Files[0].FileName);

			var headers = context.Request.Headers;
			var realName = headers["X-FileName"];
			realName = string.IsNullOrEmpty(realName)
				? Guid.NewGuid() + Path.GetExtension(context.Request.Files[0].FileName)
				: realName;

			var fullName = StorageRoot + Path.GetFileName(realName);
			var contentRange = headers["Content-Range"].Replace("bytes ", string.Empty);
			var size = long.Parse(contentRange.Substring(contentRange.IndexOf("/", StringComparison.CurrentCulture) + 1));
			var startposition = long.Parse(contentRange.Substring(0, contentRange.IndexOf("-", StringComparison.CurrentCulture)));

			using (var fs = new FileStream(fullName, FileMode.OpenOrCreate, FileAccess.Write))
			{
				fs.SetLength(size);
				fs.Seek(startposition, SeekOrigin.Begin);

				var buffer = new byte[1024];

				var l = inputStream.Read(buffer, 0, 1024);
				while (l > 0)
				{
					fs.Write(buffer, 0, l);
					l = inputStream.Read(buffer, 0, 1024);
				}
				fs.Flush();
				fs.Close();
			}
			statuses.Add(new FilesStatus(fileName, new FileInfo(fullName)));
		}

		// Upload entire file
		private void UploadWholeFile(HttpContext context, List<FilesStatus> statuses)
		{
			for (var i = 0; i < context.Request.Files.Count; i++)
			{
				var file = context.Request.Files[i];

				var realFileName = Guid.NewGuid() + Path.GetExtension(file.FileName);

				var fullPath = StorageRoot + realFileName;

				file.SaveAs(fullPath);

				var fullName = Path.GetFileName(file.FileName);
				statuses.Add(new FilesStatus(fullName, realFileName, file.ContentLength, fullPath));
			}
		}

		private void WriteJsonIframeSafe(HttpContext context, List<FilesStatus> statuses)
		{
			context.Response.AddHeader("Vary", "Accept");
			try
			{
				context.Response.ContentType = context.Request["HTTP_ACCEPT"].Contains("application/json") ? "application/json" : "text/plain";
			}
			catch
			{
				context.Response.ContentType = "text/plain";
			}

			var jsonObj = _js.Serialize(statuses.ToArray());
			context.Response.Write(jsonObj);
		}

		private static bool RemoteFilename(HttpContext context)
		{
			return !string.IsNullOrEmpty(context.Request["r"]);
		}

		private static bool GivenFilename(HttpContext context)
		{
			return !string.IsNullOrEmpty(context.Request["f"]);
		}

		public static void CopyStreamTo(Stream src, Stream dest)
		{
			int num = src.CanSeek ? Math.Min((int)(src.Length - src.Position), 8192) : 8192;
			byte[] array = new byte[num];
			int num2;
			do
			{
				num2 = src.Read(array, 0, array.Length);
				dest.Write(array, 0, num2);
			}
			while (num2 != 0);
		}

		private void DeliverAisoFile(HttpContext context)
		{
			var filename = context.Request["f"];
			var filetitle = context.Request["t"];
			filetitle = string.IsNullOrEmpty(filetitle) ? filename : filetitle;
			var req = WebRequest.Create(WebConfigurationManager.AppSettings["AisoUrl"] + "/Upload.ashx?r=1&f=" + filename);
			using (var resp = req.GetResponse())
			{
				using (var stream = resp.GetResponseStream())
				{
					var buffer = new byte[1000000];
					var bytesRead = 0;
					context.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + filetitle + "\"");
					context.Response.ContentType = "application/octet-stream";
					context.Response.ClearContent();
					while ((bytesRead = stream?.Read(buffer, 0, buffer.Length) ?? 0) != 0)
					{
						context.Response.OutputStream.Write(buffer, 0, bytesRead);
					}
				}
			}
		}

		private void DeliverChedFile(HttpContext context)
		{
			var filename = context.Request["f"];
			var filetitle = context.Request["t"];
			filetitle = string.IsNullOrEmpty(filetitle) ? filename : filetitle;

			var data = DeliverChedFileStatic(context, filename, filetitle);

			context.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + filetitle + "\"");
			context.Response.ContentType = "application/octet-stream";
			context.Response.ClearContent();
			context.Response.BinaryWrite(data);
		}

		internal static byte[] DeliverChedFileStatic(HttpContext context, string filename, string filetitle)
		{
			using (var csClient = new CustomWebServiceImplClient())
			{
				if (csClient.ClientCredentials != null)
				{
					csClient.ClientCredentials.UserName.UserName = ConfigurationManager.AppSettings["CshedLogin"];
					csClient.ClientCredentials.UserName.Password = ConfigurationManager.AppSettings["CshedPass"];
				}

				return csClient.GetDocumentData(new GetDocumentRequest { DocumentId = filename });
			}
		}

		private void DeliverFile(HttpContext context)
		{
			var filename = context.Request[UploadHandlerArgs.FILE_NAME];
			var filetitle = context.Request[UploadHandlerArgs.FILE_TITLE];
			UploadHandlerArgs.FileTypeEnum fileType = (UploadHandlerArgs.FileTypeEnum)Convert.ToInt32(context.Request[UploadHandlerArgs.FILE_TYPE] ?? ((int)UploadHandlerArgs.FileTypeEnum.Others).ToString());

			filetitle = string.IsNullOrEmpty(filetitle) ? filename : filetitle;
			var filePath = StorageRoot + filename;

			if (!File.Exists(filePath))
			{
				filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Content/images", "NoImage.png");
			}
			context.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + filetitle + "\"");
			context.Response.ContentType = "application/octet-stream";
			context.Response.ClearContent();

			switch (fileType)
			{
				case UploadHandlerArgs.FileTypeEnum.Img: //если картинка, то нужно проверить нужна ли склейка по параметрам хромой глухой и т.д.
					bool deaf = bool.Parse(context.Request[UploadHandlerArgs.IMG_DEAF] ?? false.ToString());
					bool lame = bool.Parse(context.Request[UploadHandlerArgs.IMG_LAME] ?? false.ToString());
					bool blind = bool.Parse(context.Request[UploadHandlerArgs.IMG_BLIND] ?? false.ToString());
					bool accessableEnv = bool.Parse(context.Request[UploadHandlerArgs.IMG_ACCESSABLE_ENV] ?? false.ToString());

					context.Response.BinaryWrite(new AccessableEnvImageCreator(filePath, deaf, lame, blind, accessableEnv).GetImage());
					break;
				case UploadHandlerArgs.FileTypeEnum.Others: //отправялем файл по старинке
					context.Response.TransmitFile(filePath);
					break;
			}
		}


	}
}