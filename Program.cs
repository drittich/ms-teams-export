// See https://aka.ms/new-console-template for more information
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

using teams_export;

var token = @"eyJ0eXAiOiJKV1QiLCJub25jZSI6IjdleHhNa1F4aHh5eXVKZmtUcmpQeUdCaFp1RW5lT2lQTi1LQXZMdjVhNW8iLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1yNS1BVWliZkJpaTdOZDFqQmViYXhib1hXMCIsImtpZCI6Ik1yNS1BVWliZkJpaTdOZDFqQmViYXhib1hXMCJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8zMmI2Zjk4MC1mZDZlLTQ2MDQtOWUwYS0yNzM3NDBhZDAwODEvIiwiaWF0IjoxNjQxMzI2MDIzLCJuYmYiOjE2NDEzMjYwMjMsImV4cCI6MTY0MTMzMDAzOCwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkUyWmdZTmpCeG5tRGlaZjU3SGVqZ2pmcGN4V3pJdDJrV2ZmSk5hdXpxVXlUTldYbnJnVUEiLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIEV4cGxvcmVyIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IlJpdHRpY2giLCJnaXZlbl9uYW1lIjoiRFx1MDAyN0FyY3kiLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiIxMDguMTYyLjEzMS4yMjEiLCJuYW1lIjoiRFx1MDAyN0FyY3kgUml0dGljaCIsIm9pZCI6IjNjNzNkYmY1LWU1ZTAtNDdiNC05NDY4LTVlMzU2N2ZiNDRlZSIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzN0ZGRUFEQkQxNkI2IiwicmgiOiIwLkFTa0FnUG0yTW03OUJFYWVDaWMzUUswQWdiWElpOTc1MmJGSXFLMjNTTnB5VUdRcEFLQS4iLCJzY3AiOiJBdWRpdExvZy5SZWFkLkFsbCBDaGF0LlJlYWQgRGlyZWN0b3J5LlJlYWQuQWxsIG9wZW5pZCBwcm9maWxlIFVzZXIuUmVhZCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6Im4zeUdhNmwtV1J1ZFdadjc0dV9Ma3FyaUlsaENKTC1yT1hBQUplOUR5b3ciLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiIzMmI2Zjk4MC1mZDZlLTQ2MDQtOWUwYS0yNzM3NDBhZDAwODEiLCJ1bmlxdWVfbmFtZSI6ImRyaXR0aWNoQHNlbnNlaWxhYnMuY29tIiwidXBuIjoiZHJpdHRpY2hAc2Vuc2VpbGFicy5jb20iLCJ1dGkiOiI2dGtKanIxektVMm5iemdRSUEtQUFnIiwidmVyIjoiMS4wIiwid2lkcyI6WyJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXSwieG1zX3N0Ijp7InN1YiI6ImZLbGxCN3ZweUVRSGZJS2htYmZCZkU0MjNGbXZRdjdOeWU0bXREbW1BRVUifSwieG1zX3RjZHQiOjE0OTU2MTkzOTZ9.FSqhlD8sWt5LerFH0kx6XCx9y9j3velGXvqTWh3RD5fJBPBxxr888sV8qup1MUuZh_yl8TCkOI4u6Pbeiv0Z1Jp_JcxQf2nvexb9GSlWX_mVMl1iacDT-GqI04gNw323hZNeaM1vpFY_lBYZkfsqg5xXId2jDj9wE4HWiHdNfwxH8AcgBnqwtB1PjFMpyRrQrkcl7G8O6tETJbk6I6TFbhAUw_R_p8ajuSPQh4Oi2quOyjB2aJWLg04dqMKDh4QBOKDzAP4BewuYbLGj1PBuip_RjwDuEhzKVvzHUzrqTKWYyf0th3FaGHNZIYpbeqB6E2KtXjXLQwjkUyJKNpylqQ";

if (string.IsNullOrWhiteSpace(token))
	throw new Exception("You need to populate token. Go here, https://developer.microsoft.com/en-us/graph/graph-explorer, select Chat.Read permission, then copy token from test rewquest");

var targetPath = @"c:\Temp";
var targetFile = Path.Combine(targetPath, "messages.html");
var imagesSubfolder = "images";
var targetImageFolder = Path.Combine(targetPath, imagesSubfolder);

Console.WriteLine("Start");

System.IO.Directory.CreateDirectory(targetImageFolder);

var http = new HttpClient();
http.DefaultRequestHeaders.Add("Accept", "application/json");
http.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

var http2 = new HttpClient();
http2.DefaultRequestHeaders.Add("Accept", "image/png");
http2.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

var messages = new List<MessageSummary>();
MessagesResponse result = null;
string url = "https://graph.microsoft.com/beta/chats/19:20b2cb3d-54d7-4513-ac20-4716c2424d96_3c73dbf5-e5e0-47b4-9468-5e3567fb44ee@unq.gbl.spaces/messages";
int ctr = 0;
var re = new Regex(@"src=""(https:\/\/graph.microsoft.com.*?)""", RegexOptions.Compiled);
int imageCtr = 0;
do
{
	ctr++;
	var resp = http.GetAsync(url).Result;
	if (!resp.IsSuccessStatusCode)
	{
		Console.WriteLine("Request failed, exiting");
		Console.WriteLine("Status code:" + resp.StatusCode.ToString());
		Environment.Exit(1);
	}
	var resultJson = resp.Content.ReadAsStringAsync().Result;
	result = System.Text.Json.JsonSerializer.Deserialize<MessagesResponse>(resultJson);
	url = string.IsNullOrWhiteSpace(result.odatanextLink) ? null : result.odatanextLink;
	Console.WriteLine("Page " + ctr);

	foreach (var message in result.value)
	{
		if (message.from == null)
			continue;

		//if (message.body.content.Contains("can you split that out as csv or something"))
		//	Debugger.Break();

		var isHtml = message.body.contentType == "html";
		var massagedContent = isHtml ? message.body.content : HttpUtility.HtmlEncode(message.body.content);
		if (isHtml)
		{

			MatchCollection matches = re.Matches(massagedContent);
			foreach (Match match in matches)
			{
				imageCtr++;
				var imageUrl = match.Groups[1].Value;
				Console.WriteLine($"Found image: {imageUrl}");
				var relativewImagePath = GetImage(imageUrl, imageCtr);
				massagedContent = massagedContent.Replace(imageUrl, relativewImagePath);
			}
		}
		var summary = new MessageSummary() { Body = massagedContent, From = message.from.user.displayName, Date = message.createdDateTime };
		messages.Add(summary);
	}
}
while (url != null);

string GetImage(string imageUrl, int imageNum)
{
	var img = http2.GetByteArrayAsync(imageUrl).Result;
	var imageFileName = "image_" + imageNum + ".png";
	using (var ms = new MemoryStream(img))
	{
		var savePath = Path.Combine(targetImageFolder, imageFileName);
		using (var fs = new FileStream(savePath, FileMode.Create))
			ms.WriteTo(fs);
	}
	return imagesSubfolder + "/" + imageFileName;
}

var htmlTemplate = @"
<!DOCTYPE html>
<html lang=""en"">
	<head>
		<meta charset=""UTF-8"">
		<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">
		<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
		<title>Document</title>
		<style>
			body {
				font-family: system, -apple-system, "".SFNSText-Regular"", ""San Francisco"", ""Roboto"", ""Segoe UI"", ""Helvetica Neue"", ""Lucida Grande"", sans-serif;
}
			.message {
				margin: 0px 0px 12px 0px;
			}
			.message-from {
				margin: 0px 0px 0px 0px;
				font-size: 80%;
			}
			.message-body {
				margin: 0px 0px 0px 0px;
			}	
		</style>
	</head>
	<body>
	{{BODY}}	
	</body>
</html>";

var html = new StringBuilder();
foreach (var message in messages.OrderBy(m => m.Date))
{
	html.AppendLine($"<div class=\"message\">");
	html.AppendLine($"  <div class=\"message-from\">{message.From} {message.Date.ToString("yyyy-MM-dd hh:mm:ss tt")}</div>");
	html.AppendLine($"  <div class=\"message-body\">{message.Body}</div>");
	html.AppendLine($"</div>");
}
htmlTemplate = htmlTemplate.Replace("{{BODY}}", html.ToString());
System.IO.File.WriteAllText(targetFile, htmlTemplate.ToString());

Console.WriteLine($"Found {imageCtr} images");

Console.WriteLine("Done");
