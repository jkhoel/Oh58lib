using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System.Resources;
using System.Runtime.InteropServices;



namespace Excel.Oh58lib.Ribbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        private Application _excel;
        private IRibbonUI _thisRibbon;

        public override string GetCustomUI(string RibbonID)
        {
            _excel = ExcelDnaUtil.Application;

            var ribbonXml = GetCustomRibbonXML("RibbonController");
            return ribbonXml;
        }

        #region Initialization and Disposing

        private string GetCustomRibbonXML(string customRibbonXMLFileName)
        {
            string ribbonXml;
            var thisAssembly = typeof(RibbonController).Assembly;
            var resourceName = typeof(RibbonController).Namespace + $".{customRibbonXMLFileName}.xml";

            using (Stream stream = thisAssembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                ribbonXml = reader.ReadToEnd();
            }

            if (ribbonXml == null)
            {
                throw new MissingManifestResourceException(resourceName);
            }
            return ribbonXml;
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            if (ribbon == null)
            {
                throw new ArgumentNullException(nameof(ribbon));
            }

            _thisRibbon = ribbon;

            _excel.WorkbookActivate += OnInvalidateRibbon;
            _excel.WorkbookDeactivate += OnInvalidateRibbon;
            _excel.SheetActivate += OnInvalidateRibbon;
            _excel.SheetDeactivate += OnInvalidateRibbon;

            if (_excel.ActiveWorkbook == null)
            {
                _excel.Workbooks.Add();
            }
        }

        private void OnInvalidateRibbon(object obj)
        {
            _thisRibbon.Invalidate();
        }

        #endregion

        public void OnButtonPressed(IRibbonControl control)
        {
            // TODO: Build the mission object somehow, and then export it
        }


    }
}

// https://andysprague.com/2017/02/03/my-first-custom-excel-ribbon-using-excel-dna/

//namespace Excel.Oh58lib.Ribbon
//{
//    [ComVisible(true)]
//    internal class RibbonController: ExcelRibbon
//    {
//        private Application _excel;
//        private IRibbonUI _thisRibbon;

//        public override string GetCustomUI(string ribbonId)
//        {
//            _excel = ExcelDnaUtil.Application;

//            string ribbonXml = GetCustomRibbonXML();
//            return ribbonXml;
//        }

//        #region Initialization and Updating Methods

//        private string GetCustomRibbonXML()
//        {
//            string ribbonXml;
//            var thisAssembly = typeof(RibbonController).Assembly;
//            var resourceName = typeof(RibbonController).Namespace + ".CustomRibbon.xml";

//            using (Stream stream = thisAssembly.GetManifestResourceStream(resourceName))
//            using (StreamReader reader = new StreamReader(stream))
//            {
//                ribbonXml = reader.ReadToEnd();
//            }

//            if (ribbonXml == null)
//            {
//                throw new MissingManifestResourceException(resourceName);
//            }
//            return ribbonXml;
//        }

//        public void OnLoad(IRibbonUI ribbon)
//        {
//            if (ribbon == null)
//            {
//                throw new ArgumentNullException(nameof(ribbon));
//            }

//            _thisRibbon = ribbon;

//            _excel.WorkbookActivate += OnInvalidateRibbon;
//            _excel.WorkbookDeactivate += OnInvalidateRibbon;
//            _excel.SheetActivate += OnInvalidateRibbon;
//            _excel.SheetDeactivate += OnInvalidateRibbon;

//            if (_excel.ActiveWorkbook == null)
//            {
//                _excel.Workbooks.Add();
//            }
//        }

//        private void OnInvalidateRibbon(object obj)
//        {
//            _thisRibbon.Invalidate();
//        }

//        #endregion

//        #region Ribbon Commands

//        public void OnButtonPressed(IRibbonControl control)
//        {
//            //MessageBox.Show("Hello from control " + control.Id);

//        }

//        #endregion
//    }
//}
