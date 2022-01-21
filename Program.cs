using System.Text;
using System.Text.RegularExpressions;
using System.Web;

using Microsoft.Extensions.Configuration;

using teams_export;

IConfiguration config = new ConfigurationBuilder()
	.AddJsonFile("appsettings.json")
	.Build();

var settings = config.GetRequiredSection("Settings").Get<Settings>();

if (string.IsNullOrWhiteSpace(settings.Token))
	throw new Exception("You need to provide a token. Go here, https://developer.microsoft.com/en-us/graph/graph-explorer, use beta API, use endpoint https://graph.microsoft.com/beta/chats and select Chat.Read permission, then update appsettings.json accordingly");

var targetPath = @"c:\Temp";
var targetFile = Path.Combine(targetPath, "messages.html");
var imagesSubfolder = "images";
var targetImageFolder = Path.Combine(targetPath, imagesSubfolder);

Console.WriteLine("Start");

System.IO.Directory.CreateDirectory(targetImageFolder);

var http = new HttpClient();
http.DefaultRequestHeaders.Add("Accept", "application/json");
http.DefaultRequestHeaders.Add("Authorization", "Bearer " + settings.Token);

var http2 = new HttpClient();
http2.DefaultRequestHeaders.Add("Accept", "image/png");
http2.DefaultRequestHeaders.Add("Authorization", "Bearer " + settings.Token);

var messages = new List<MessageSummary>();
MessagesResponse result = null;
// go to https://teams.microsoft.com/ and grab the ID for the conversation you want
// it will look something like: 19:20b2cb3d-54d7-4513-ac20-4716c2424d96_3c73dbf5-e5e0-47b4-9468-5e3567fb44ee@unq.gbl.spaces
string url = $"https://graph.microsoft.com/beta/chats/{settings.MessageID}/messages";
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
		var summary = new MessageSummary() { Body = massagedContent, From = message.from?.user?.displayName ?? "(unknown)", Date = message.createdDateTime };
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
