using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Xml;

namespace ZIPApp
{
	public class OpenXmlWriter
	{
		XmlWriter writer = null;

		public OpenXmlWriter(XmlWriter writer)
		{
			this
				.writer = writer;
		}

		public void OpenWorksheet()
		{
			writer
				.WriteStartElement("worksheet",
				"http://schemas.openxmlformats.org/spreadsheetml/2006/main");

			writer
				.WriteAttributeString("xmlns",
				"r",
				null,
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships");

			writer
				.WriteAttributeString("xmlns",
				"mc",
				null,
				"http://schemas.openxmlformats.org/markup-compatibility/2006");

			writer
				.WriteAttributeString("Ignorable",
				"http://schemas.openxmlformats.org/markup-compatibility/2006",
				"x14ac xr xr2 xr3");

			writer
				.WriteAttributeString("xmlns",
				"x14ac",
				null,
				"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

			writer
				.WriteAttributeString("xmlns",
				"xr",
				null,
				"http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

			writer
				.WriteAttributeString("xmlns",
				"xr2",
				null,
				"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");

			writer
				.WriteAttributeString("xmlns",
				"xr3",
				null,
				"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");

			writer
				.WriteAttributeString("uid",
				"http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
				String
				.Format("{{{0}}}",
				Guid
				.NewGuid()
				.ToString()));

			writer
				.WriteStartElement("sheetData");
		}

		public void CloseWorksheet()
		{
			writer
				.WriteEndElement();

			writer
				.WriteEndElement();
		}

		public void OpenRow()
		{
			writer
				.WriteStartElement("row");
		}

		public void CloseRow()
		{
			writer
				.WriteEndElement();
		}

		public void WriteValue(object value)
		{
			writer
				.WriteStartElement("c");

			writer
				.WriteElementString("v",
				String
				.Format("{0}",
				value));

			writer
				.WriteEndElement();
		}

		public void WriteNumber(object value)
		{
			writer
				.WriteStartElement("c");

			writer
				.WriteStartElement("v");

			writer
				.WriteValue(value);

			writer
				.WriteEndElement();

			writer
				.WriteEndElement();
		}

		public void WriteString(string value)
		{
			writer
				.WriteStartElement("c");

			writer
				.WriteAttributeString("t",
				"inlineStr");

			writer
				.WriteStartElement("is");

			writer
				.WriteElementString("t",
				value);

			writer
				.WriteEndElement();

			writer
				.WriteEndElement();
		}
	}
}