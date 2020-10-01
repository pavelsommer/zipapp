<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="ZIPApp.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title></title>
</head>
<body>
	<form id="form1" runat="server">
		<div>
			<asp:Button ID="button1" runat="server" Text="Save to response" OnClick="button1_Click" />
		</div>
		<div>
			<asp:Button ID="button2" runat="server" Text="Save to file" OnClick="button2_Click" />
		</div>
	</form>
</body>
</html>
