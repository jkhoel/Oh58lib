using System.Xml.Serialization;

namespace Excel.Oh58lib
{
	internal class Utils
	{
		public static T DeserializeFromXmlFile<T>(string filePath)
		{
			XmlSerializer serializer = new XmlSerializer(typeof(T));
			using (StreamReader reader = new StreamReader(filePath))
			{
                return (T)serializer.Deserialize(reader);
			}
		}
	}
}
