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