using ExcelDna.Integration;
using CoordinateSharp;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace Excel.Oh58lib
{
	[ComVisible(true)]
	public class Geo
	{
		private static readonly Theaters Theaters = new();

		private readonly Theater _activeTheater;

		public Geo()
		{
			_activeTheater = Theaters.Kola;
		}

		#region EXCEL FUNCTIONS

		[ExcelFunction(Description = "Convert Latitude and Longitude in DMS format to MGRS")]
		public static string MGRS(string coordinate)
		{
			if (string.IsNullOrWhiteSpace(coordinate))
				return string.Empty;

			var geo = new Geo();
			return geo.LatLonToMGRS(coordinate);
		}

		[ExcelFunction(Description = "Convert Latitude and Longitude in DMS format to DCS coordiantes")]
		public static string DcsCoordinates(string coordinate)
		{
			if (string.IsNullOrWhiteSpace(coordinate))
				return string.Empty;

			var geo = new Geo();

			var coords = geo.ToDcsCoordiantes(coordinate);

			return $"X{coords.Easting} Z{coords.Northing}";
		}

		[ExcelFunction(Description = "Convert Latitude and Longitude in DMS format to DCS coordiantes and returns Easting (X)")]
		public static string DcsCoordinateX(string coordinate)
		{
			if (string.IsNullOrWhiteSpace(coordinate))
				return string.Empty;

			var geo = new Geo();

			var coords = geo.ToDcsCoordiantes(coordinate);

			return $"{coords.Easting}";
		}

		[ExcelFunction(Description = "Convert Latitude and Longitude in DMS format to DCS coordiantes and returns Northing (Z)")]
		public static string DcsCoordinateZ(string coordinate)
		{
			if (string.IsNullOrWhiteSpace(coordinate))
				return string.Empty;

			var geo = new Geo();

			var coords = geo.ToDcsCoordiantes(coordinate);

			return $"{coords.Northing}";
		}

		#endregion

		#region C# FUNCTIONS
		/// <summary>
		/// Converts a coordinate string formatted as "N67:15:10 E014:55:55" to MGRS
		/// </summary>
		/// <param name="coordinate"></param>
		/// <returns></returns>
		public string LatLonToMGRS(string coordinate)
		{
			var (latitude, longitude) = ParseCoordinate(coordinate);
			Coordinate c = new Coordinate(latitude, longitude);
			string mgrs = c.MGRS.ToString();
			return mgrs;
		}

		/// <summary>
		/// Converts a coordinate string formatted as "N67:15:10 E014:55:55" to DCS coordinates
		/// </summary>
		/// <param name="coordinate"></param>
		/// <returns></returns>
		public (double Northing, double Easting, int Zone) ToDcsCoordiantes(string coordinate)
		{
			var (latitude, longitude) = ParseCoordinate(coordinate);

			return UtmConverter.FromLatLon(latitude, longitude, _activeTheater.UTM_zone, _activeTheater.DCS_origin_easting, _activeTheater.DCS_origin_northing);
		}


		/// <summary>
		/// Parses a coordinate string formatted as "N67:15:10 E014:55:55"
		/// </summary>
		/// <param name="coordinate"></param>
		/// <returns>A tuple of latitude and logitude as doubles</returns>
		/// <exception cref="ArgumentException"></exception>
		private (double latitude, double longitude) ParseCoordinate(string coordinate)
		{
			var dmsRegex = new Regex(@"([NS])\s*(\d+):(\d+):(\d+(?:\.\d+)?)\s*([EW])\s*(\d+):(\d+):(\d+(?:\.\d+)?)");
			var ddmRegex = new Regex(@"([NS])\s*(\d+):(\d+(?:\.\d+)?)\s*([EW])\s*(\d+):(\d+(?:\.\d+)?)");
			var ddmAltRegex = new Regex(@"([NS])\s*(\d+)\s+(\d+(?:\.\d+)?)\s*([EW])\s*(\d+)\s+(\d+(?:\.\d+)?)");

			Match match;
			if ((match = dmsRegex.Match(coordinate)).Success)
			{
				string latitudePart = $"{match.Groups[1].Value}{match.Groups[2].Value}:{match.Groups[3].Value}:{match.Groups[4].Value}";
				string longitudePart = $"{match.Groups[5].Value}{match.Groups[6].Value}:{match.Groups[7].Value}:{match.Groups[8].Value}";

				double latitude = ParseDMS(latitudePart);
				double longitude = ParseDMS(longitudePart);

				return (latitude, longitude);
			}
			else if ((match = ddmRegex.Match(coordinate)).Success)
			{
				string latitudePart = $"{match.Groups[1].Value}{match.Groups[2].Value}:{match.Groups[3].Value}";
				string longitudePart = $"{match.Groups[4].Value}{match.Groups[5].Value}:{match.Groups[6].Value}";

				double latitude = ParseDMS(latitudePart);
				double longitude = ParseDMS(longitudePart);

				return (latitude, longitude);
			}
			else if ((match = ddmAltRegex.Match(coordinate)).Success)
			{
				string latitudePart = $"{match.Groups[1].Value}{match.Groups[2].Value} {match.Groups[3].Value}";
				string longitudePart = $"{match.Groups[4].Value}{match.Groups[5].Value} {match.Groups[6].Value}";

				double latitude = ParseDMS(latitudePart);
				double longitude = ParseDMS(longitudePart);

				return (latitude, longitude);
			}
			else
			{
				throw new ArgumentException("Invalid coordinate format");
			}
		}

		/// <summary>
		/// Handles parsing of DMS coordinate parts
		/// </summary>
		/// <param name="dms"></param>
		/// <returns>Decimal Degrees as a Double</returns>
		/// <exception cref="ArgumentException"></exception>
		private double ParseDMS(string dms)
		{
			char direction = dms[0];
			string coordinatePart = dms.Substring(1).Trim();
			double decimalDegrees;

			var dmsRegex = new Regex(@"(\d+):(\d+):(\d+(?:\.\d+)?)");
			var ddmRegex = new Regex(@"(\d+):(\d+(?:\.\d+)?)");
			var ddmAltRegex = new Regex(@"(\d+)\s+(\d+(?:\.\d+)?)");

			if (dmsRegex.IsMatch(coordinatePart))
			{
				// DMS format
				var match = dmsRegex.Match(coordinatePart);
				double degrees = double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
				double minutes = double.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
				double seconds = double.Parse(match.Groups[3].Value, System.Globalization.CultureInfo.InvariantCulture);

				decimalDegrees = degrees + (minutes / 60) + (seconds / 3600);
			}
			else if (ddmRegex.IsMatch(coordinatePart))
			{
				// DDM format
				var match = ddmRegex.Match(coordinatePart);
				double degrees = double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
				double decimalMinutes = double.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);

				decimalDegrees = degrees + (decimalMinutes / 60);
			}
			else if (ddmAltRegex.IsMatch(coordinatePart))
			{
				// Alternative DDM format
				var match = ddmAltRegex.Match(coordinatePart);
				double degrees = double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
				double decimalMinutes = double.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);

				decimalDegrees = degrees + (decimalMinutes / 60);
			}
			else
			{
				throw new ArgumentException("Invalid coordinate format");
			}

			if (direction == 'S' || direction == 'W')
				decimalDegrees = -decimalDegrees;

			return decimalDegrees;
		}
	}

	#endregion

}
