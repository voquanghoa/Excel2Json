using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Json2
{
	[DataContract]
	public class Question
	{
		[DataMember(Name = "title")]
		public string Title
		{
			get; set;
		}

		[DataMember(Name = "answer")]
		public string Answer
		{
			get; set;
		}
	}

	[DataContract]
	public class Level
	{
		[DataMember(Name = "name")]
		public string Name
		{
			get; set;
		}

		[DataMember(Name = "questions")]
		public List<Question> Questions
		{
			get; set;
		}

		public Level(string name)
		{
			Name = name;
			Questions = new List<Question>();
		}
	}

	[DataContract]
	public class Data : List<Level>
	{

	}
}
