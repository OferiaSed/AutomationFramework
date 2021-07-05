using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AutomationFramework
{
    public class clsData
    {
        //Variables
        private SLDocument _objFile;
        private int _intCurrentRow;
        private bool _blColumnNames;
        private bool _blSheetInUse;
        private int _intColCount;
        private int _intRowCount;
        private int _intColStartIndex;
        private int _intRowStartIndex;
        private Dictionary<string, int> _dicHeader;

        //Methos
        public int CurrentRow
        {
            get { return _intCurrentRow; }
            set { _intCurrentRow = (value > _intRowStartIndex ? value : _intRowStartIndex); }
        }

        public bool HasColumnNames
        {
            get { return _blColumnNames; }
            set
            {
                _blColumnNames = value;
                _intRowStartIndex = (value ? 2 : 1);
            }
        }

        public int ColumnCount
        {
            get { return _intColCount; }
        }

        public int RowCount
        {
            get { return _intRowCount; }
        }

        private int RowStartIndex
        {
            get { return _intRowStartIndex; }
            set
            {
                _intRowStartIndex = value;
                _intCurrentRow = (_intCurrentRow > _intRowStartIndex ? _intCurrentRow : _intRowStartIndex);
            }
        }

        private int ColumnStartIndex
        {
            get { return _intColStartIndex; }
            set { _intColStartIndex = value; }
        }

        //Constructor
        public clsData()
        {
            _intCurrentRow = 2;
            _blColumnNames = true;
            _blSheetInUse = true;
            _intColCount = -1;
            _intRowCount = -1;
            _intRowStartIndex = 2;
            _intColStartIndex = 1;
        }

        //Private Methods
        private void _GetRowCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intRowCount = _objFile.GetWorksheetStatistics().NumberOfRows > -1 ? _objFile.GetWorksheetStatistics().NumberOfRows : 0;
            }
        }

        private void _GetColumCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intColCount = _objFile.GetWorksheetStatistics().NumberOfColumns > -1 ? _objFile.GetWorksheetStatistics().NumberOfColumns : 0;
            }
        }

        public void fnLoadFile(string pstrFilePath, string pstrSheet)
        {
            if (!string.IsNullOrEmpty(pstrFilePath) && File.Exists(pstrFilePath))
            {
                _objFile = new SLDocument(pstrFilePath);
                if (!string.IsNullOrEmpty(pstrSheet))
                {
                    if (_objFile.GetSheetNames().Contains(pstrSheet)) { _objFile.SelectWorksheet(pstrSheet); }
                }
                _blColumnNames = true;
                _blSheetInUse = true;
                _intCurrentRow = 2;
                _intRowStartIndex = _objFile.GetWorksheetStatistics().StartRowIndex;
                _intColStartIndex = _objFile.GetWorksheetStatistics().StartColumnIndex;
                _dicHeader = _GetHeaders(_objFile);
                _GetRowCount();
                _GetColumCount();

            }
        }

        private string fnRemoveEspecialChar(string strValue)
        {
            Regex r = new Regex("(?:[^a-z0-9 ]|(?<=['\"])s)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
            return r.Replace(strValue, String.Empty);
        }

        private Dictionary<string, int> _GetHeaders(SLDocument pobjDoc)
        {
            if (pobjDoc != null)
            {
                Dictionary<string, int> dicHeaders = new Dictionary<string, int>();
                for (int cols = 1; cols <= pobjDoc.GetWorksheetStatistics().NumberOfColumns; cols++)
                {
                    if (pobjDoc.GetCellValueAsString(1, cols) != "")
                    {
                        dicHeaders.Add(fnRemoveEspecialChar(pobjDoc.GetCellValueAsString(1, cols)), cols);
                    }
                    else
                    {
                        break;
                    }
                }
                return dicHeaders;
            }
            else
            {
                return null;
            }
        }

        public string fnGetValue(string pstrColumnName, [Optional] string pstrDefaultValue)
        {
            if (string.IsNullOrEmpty(pstrDefaultValue)) { pstrDefaultValue = string.Empty; }
            if (_objFile != null && _blSheetInUse && _dicHeader.Count > 0)
            {
                if (_dicHeader.ContainsKey(pstrColumnName))
                {
                    pstrDefaultValue = _objFile.GetCellValueAsString(_intCurrentRow, _dicHeader[pstrColumnName]);
                }
                else
                {
                    pstrDefaultValue = "";
                }
            }
            return pstrDefaultValue;
        }

    }
}
