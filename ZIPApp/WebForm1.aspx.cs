using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.IO.Compression;

using ICSharpCode.SharpZipLib.Zip;

namespace ZIPApp
{
	public partial class WebForm1 : System.Web.UI.Page
	{
		protected void Page_Load(object sender, EventArgs e)
		{
		}

		// this produces "corrupted" xlsx
		protected void button1_Click(object sender, EventArgs e)
		{
			string fileName = @"output.xlsx";

			Response
				.Clear();

			Response
				.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

			Response
				.AppendHeader("content-disposition",
				String
				.Format("attachment; filename=\"{0}\";filename*=utf-8''{1}",
				fileName,
				HttpUtility
				.UrlPathEncode(fileName)));

			Response
				.CacheControl = "Private";

			Response
				.BufferOutput = false;

			ComposeOutput(Response
				.OutputStream);

			Response
				.End();
		}

		// this produces perfectly valid xlsx
		protected void button2_Click(object sender, EventArgs e)
		{
			string fileName = @"output.xlsx";

			using (var outputStream = new FileStream(Server
				.MapPath(String
				.Format("~/App_Data/{0}",
				fileName)),
				FileMode
				.Create))
			{
				ComposeOutput(outputStream);
			}
		}

		protected void ComposeOutput(Stream outputStream)
		{
			using (ZipOutputStream zipStream = new ZipOutputStream(outputStream))
			{
				zipStream
					.IsStreamOwner = true;

				ZipEntry zipEntry = new ZipEntry(ZipEntry
					.CleanName(@"xl\worksheets\sheet1.xml"));

				zipStream
					.PutNextEntry(zipEntry);

				// write dynamic content to output xlsx
				XmlWriter xmlWriter = XmlWriter
					.Create(zipStream);

				var excelWriter = new OpenXmlWriter(xmlWriter);

				excelWriter
					.OpenWorksheet();

				for (int i = 0; i < 1000; i++)
				{
					excelWriter
						.OpenRow();

					excelWriter
						.WriteString(String
						.Format("Hi there {0}",
						i));

					excelWriter
						.CloseRow();

					zipStream
						.Flush();
				}

				excelWriter
					.CloseWorksheet();

				xmlWriter
					.Close();

				// writes rest of xlsx structure ([Content_Types].xml, workbook.xml etc.)
				WriteFiles(zipStream,
					Server
					.MapPath("~/App_Data/Excel"));

				zipStream
					.Finish();

				zipStream
					.Flush();

				zipStream
					.Close();
			}
		}

		void WriteFiles(ZipOutputStream zipStream, string path, params string[] exclude)
		{
			foreach (var filePath in Directory
				.GetFiles(path,
				"*.*",
				SearchOption
				.AllDirectories))
			{
				var relativePath = filePath
					.Replace(String
					.Format("{0}\\",
					path),
					null);

				if (exclude
					.Contains(relativePath))
					continue;

				WriteFile(zipStream,
					relativePath,
					filePath);
			}
		}

		void WriteFile(ZipOutputStream zipStream, string entryPath, string filePath)
		{
			ZipEntry zipEntry = new ZipEntry(ZipEntry
				.CleanName(entryPath));

			zipStream
				.PutNextEntry(zipEntry);

			byte[] buffer = File
				.ReadAllBytes(filePath);

			zipStream
				.Write(buffer,
				0,
				buffer
				.Length);

			zipStream
				.CloseEntry();
		}
	}
}