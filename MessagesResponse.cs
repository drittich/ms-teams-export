using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace teams_export
{

	public class MessagesResponse
	{
		[JsonPropertyName("@odata.context")]
		public string odatacontext { get; set; }
		[JsonPropertyName("@odata.count")]
		public int odatacount { get; set; }
		[JsonPropertyName("@odata.nextLink")]
		public string odatanextLink { get; set; }
		public Value[] value { get; set; }
	}

	public class Value
	{
		public string id { get; set; }
		public object replyToId { get; set; }
		public string etag { get; set; }
		public string messageType { get; set; }
		public DateTime createdDateTime { get; set; }
		public DateTime lastModifiedDateTime { get; set; }
		public DateTime? lastEditedDateTime { get; set; }
		public object deletedDateTime { get; set; }
		public object subject { get; set; }
		public object summary { get; set; }
		public string chatId { get; set; }
		public string importance { get; set; }
		public string locale { get; set; }
		public object webUrl { get; set; }
		public object channelIdentity { get; set; }
		public object policyViolation { get; set; }
		public object eventDetail { get; set; }
		public From from { get; set; }
		public Body body { get; set; }
		public Attachment[] attachments { get; set; }
		public object[] mentions { get; set; }
		public Reaction[] reactions { get; set; }
	}

	public class From
	{
		public object application { get; set; }
		public object device { get; set; }
		public User user { get; set; }
	}

	public class User
	{
		public string id { get; set; }
		public string displayName { get; set; }
		public string userIdentityType { get; set; }
	}

	public class Body
	{
		public string contentType { get; set; }
		public string content { get; set; }
	}

	public class Attachment
	{
		public string id { get; set; }
		public string contentType { get; set; }
		public string contentUrl { get; set; }
		public object content { get; set; }
		public string name { get; set; }
		public object thumbnailUrl { get; set; }
	}

	public class Reaction
	{
		public string reactionType { get; set; }
		public DateTime createdDateTime { get; set; }
		public User1 user { get; set; }
	}

	public class User1
	{
		public object application { get; set; }
		public object device { get; set; }
		public User2 user { get; set; }
	}

	public class User2
	{
		public string id { get; set; }
		public object displayName { get; set; }
		public string userIdentityType { get; set; }
	}

}
