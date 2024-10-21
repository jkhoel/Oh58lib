using Excel.Oh58lib.Models;
using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace Excel.Oh58lib
{
    [ComVisible(true)]
    public class Addin
    {
        #region Properties
        private static readonly Theaters Theaters = new();


        #endregion

        #region Fields and Constructor

        private static readonly Geo _geo = new(Theaters.Kola);

        public Addin() {

        }

        #endregion

        #region EXCEL FUNCTIONS

        [ExcelFunction(Description = "Convert Latitude and Longitude in DMS format to MGRS")]
        public static string MGRS(string coordinate)
        {
            if (string.IsNullOrWhiteSpace(coordinate))
                return string.Empty;

            return _geo.LatLonToMGRS(coordinate);
        }

        [ExcelFunction(Description = "Convert Latitude and Longitude in DMS format to DCS coordiantes")]
        public static string DcsCoordinates(string coordinate)
        {
            if (string.IsNullOrWhiteSpace(coordinate))
                return string.Empty;

            var coords = _geo.ToDcsCoordiantes(coordinate);

            return $"X{coords.Easting} Z{coords.Northing}";
        }

        [ExcelFunction(Description = "Convert Latitude and Longitude in DMS format to DCS coordiantes and returns Easting (X)")]
        public static string DcsCoordinateX(string coordinate)
        {
            if (string.IsNullOrWhiteSpace(coordinate))
                return string.Empty;

            var coords = _geo.ToDcsCoordiantes(coordinate);

            return $"{coords.Easting}";
        }

        [ExcelFunction(Description = "Convert Latitude and Longitude in DMS format to DCS coordiantes and returns Northing (Z)")]
        public static string DcsCoordinateZ(string coordinate)
        {
            if (string.IsNullOrWhiteSpace(coordinate))
                return string.Empty;

            var coords = _geo.ToDcsCoordiantes(coordinate);

            return $"{coords.Northing}";
        }

        #endregion
    }
}
