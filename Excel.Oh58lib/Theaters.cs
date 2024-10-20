//using System.Xml.Serialization;

//namespace Excel.Oh58lib
//{
//	[XmlRoot("Theaters")]
//	public class Theaters
//	{
//		[XmlElement("Theater")]
//		public List<Theater>? TheaterList { get; set; }
//	}

//	public class Theater
//	{
//		public string Name { get; set; }
//		public double DCS_origin_northing { get; set; }
//		public double DCS_origin_easting { get; set; }
//		public int UTM_zone { get; set; }
//		public string Hemisphere { get; set; }
//		public int WinterTimeDelta { get; set; }
//		public int SummerTimeDelta { get; set; }
//	}
//}

namespace Excel.Oh58lib
{
	public class Theaters
	{
		public Theater Kola = new Theater
		{
			Name = "Kola",
			DCS_origin_northing = -7543624.999999979,
			DCS_origin_easting = -62702.00000000087,
			UTM_zone = 34,
			Hemisphere = "N",
			WinterTimeDelta = 1,
			SummerTimeDelta = 2
		};


		public Theater Syria = new Theater
		{
			Name = "Syria",
			DCS_origin_northing = 3879865,
			DCS_origin_easting = 217198,
			UTM_zone = 37,
			Hemisphere = "N",
			WinterTimeDelta = 3,
			SummerTimeDelta = 4
		};
	}

	public class Theater
	{
		public string? Name { get; set; }
		public double DCS_origin_northing { get; set; }
		public double DCS_origin_easting { get; set; }
		public int UTM_zone { get; set; }
		public string? Hemisphere { get; set; }
		public int WinterTimeDelta { get; set; }
		public int SummerTimeDelta { get; set; }
	}
}