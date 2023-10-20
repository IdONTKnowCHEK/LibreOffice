using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using uno;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.sheet;
using unoidl.com.sun.star.table;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.util;

namespace LOCalc
{
    public class Class1
    {
        private XComponentContext m_xComponentContext = null;
        private XMultiServiceFactory m_xMSFactory = null;
        private XComponentLoader m_xLoader = null;

        private XSpreadsheetDocument m_xDocument;
        private XSpreadsheet m_xSpreadsheet;


        public bool bOpenOfficeInstalled
        {
            get
            {
                RegistryKey RegKey = Registry.CurrentUser.OpenSubKey
                                     (@"SOFTWARE\OpenOffice", writable: false);
                return RegKey != null;
            }
        }
        public bool bStartOpenOfficeLoader()
        {
            try
            {
                if (m_xComponentContext == null)
                    m_xComponentContext = uno.util.Bootstrap.bootstrap();
                if (m_xMSFactory == null)
                    m_xMSFactory =
                   (XMultiServiceFactory)m_xComponentContext.getServiceManager();
                if (m_xLoader == null)
                    m_xLoader = (XComponentLoader)m_xMSFactory.createInstance
                               ("com.sun.star.frame.Desktop");
                return true;
            }
            catch (unoidl.com.sun.star.uno.Exception e1)
            {
                Debug.Print(e1.Message);
            }
            return false;
        }

        public bool bCreateWorkbook()
        {
            try
            {
                XComponent xComponent =
                    m_xLoader.loadComponentFromURL
                   ("private:factory/scalc", "_blank", 0, new PropertyValue[0]);

                m_xDocument = (XSpreadsheetDocument)xComponent;
                XSpreadsheets xSheets = m_xDocument.getSheets();
                XIndexAccess xIndexAccess = (XIndexAccess)xSheets;
                Any any = xIndexAccess.getByIndex(0);
                m_xSpreadsheet = (XSpreadsheet)any.Value;
                return true;
            }
            catch (unoidl.com.sun.star.uno.Exception e)
            {
                Debug.Print(e.Message);
            }
            return false;
        }

        private XCell xGetXCellByName(string strCellName)
        {
            XCellRange xCR = m_xSpreadsheet.getCellRangeByName(strCellName);
            XCell xCell = xCR.getCellByPosition(nColumn: 0, nRow: 0);
            return xCell;
        }

        public void bSetValue(string strCellName, double dValue)
        {
            XCell xCell = xGetXCellByName(strCellName);
            xCell.setValue(dValue);
        }
        public void bSetFormula(string strCellName, string strFormula)
        {
            XCell xCell = xGetXCellByName(strCellName);
            xCell.setFormula(strFormula);
        }
        public void bSetText(string strCellName, string strText)
        {
            XCell xCell = xGetXCellByName(strCellName);
            XSimpleText xST = (XSimpleText)xCell;
            XTextCursor xCursor = xST.createTextCursor();
            xST.insertString(xCursor, strText, bAbsorb: false);
        }

        public bool bSetDate(string strCellName, int nYear, int nMonth, int nDay)
        {
            // Set the date value.
            XCell xCell = xGetXCellByName(strCellName);
            System.DateTime dt = new System.DateTime(nYear, nMonth, nDay);
            string strDateStr = dt.ToString("M/dd/yyyy");

            // You can also set "text" using the setFormula method
            xCell.setFormula(strDateStr);

            // Set date format.
            XNumberFormatsSupplier xFormatsSupplier = (XNumberFormatsSupplier)m_xDocument;
            XNumberFormatTypes xFormatTypes =
                  (XNumberFormatTypes)xFormatsSupplier.getNumberFormats();
            int nFormat = xFormatTypes.getStandardFormat(NumberFormat.DATE, new Locale());

            XPropertySet xPropSet = (XPropertySet)xCell;
            xPropSet.setPropertyValue("NumberFormat", new Any(nFormat));
            return true;
        }

        public void bSetBackgroundColor(string strCellName, Color clr)
        {
            XCell xCell = xGetXCellByName(strCellName);
            XPropertySet xPropSet = (XPropertySet)xCell;
            UInt32 unClr = (UInt32)clr.R << 16 | (UInt32)clr.G << 8 | (UInt32)clr.B;
            xPropSet.setPropertyValue("CellBackColor", new Any(unClr));
        }

        public string strSaveWorkbook(string strFilePath = null)
        {
            XStorable xStorable = (XStorable)m_xDocument;
            string strFileURL = string.Empty;

            // If we don't have a file path passed in, check if the file has
            // been previously saved. If so, use that file name
            if (string.IsNullOrEmpty(strFilePath))
            {
                // If we already have a file path, use it
                if (xStorable.hasLocation())
                {
                    strFileURL = xStorable.getLocation();
                    strFilePath = strUrlToPath(strFileURL);
                }
                // Otherwise prompt for one
                else
                {
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "ODF Spreadsheet (*.ods)|*.ods";
                    DialogResult dr = sfd.ShowDialog();
                    if (dr == DialogResult.Cancel)
                        return String.Empty;
                    strFilePath = sfd.FileName;
                }
            }
            if (string.IsNullOrEmpty(strFileURL))
                strFileURL = strPathToURL(strFilePath);

            // Save with no args for default file format
            xStorable.storeAsURL(strFileURL, new PropertyValue[0]);
            return strFilePath;
        }

        private string strPathToURL(string strFilePath)
        {
            string strURL = "file:///" + strFilePath.Replace("\\", "/");
            return strURL;
        }
        private string strUrlToPath(string strFileURL)
        {
            string strFilePath = strFileURL;
            if (strFileURL.Substring(0, 8) == "file:///")
                strFilePath = strFileURL.Substring(8);
            strFilePath = strFilePath.Replace("/", "\\");
            return strFilePath;
        }

        public string strSaveWorkbookAsExcel()
        {
            XStorable xStorable = (XStorable)m_xDocument;

            if (xStorable.hasLocation() == false)
                return string.Empty;

            string strFileURL = xStorable.getLocation(),
                strFilePath = strUrlToPath(strFileURL);

            PropertyValue[] apv = new PropertyValue[1];
            apv[0] = new PropertyValue();
            apv[0].Name = "FilterName";
            apv[0].Value = new Any("MS Excel 97");
            strFilePath = Path.ChangeExtension(strFilePath, "xls");
            strFileURL = strPathToURL(strFilePath);
            xStorable.storeAsURL(strFileURL, apv);
            return strFilePath;
        }


        public void bCloseWorkbook()
        {
            if (m_xDocument != null)
            {
                XComponent xComponent = (XComponent)m_xDocument;
                xComponent.dispose();
            }
        }
    }
}
