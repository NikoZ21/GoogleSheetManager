using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Xml;

namespace GoogleSheetApi
{
    /// <summary>
    /// SHEET=1 TAB. THE CLASS REPRESENTS ONE TAB AND CONTAINST LIST OF RUNCTION TO READ.
    /// </summary>
    public class Sheet
    {
        public async Task<Dictionary<string, Dictionary<string, string>>> GetUICD()
        {
            var nonEmptyCells = await ReadFromSheet(SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE,
                SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.FORMATTEDSTRING);
            SetList(nonEmptyCells);
            SetDictionary();

            return _currentSheet;
        }

        public async Task<bool> ApendUICD(Dictionary<string, Dictionary<string, string>> dictionary, List<string> keys)
        {
            if (dictionary.Count == 0 || dictionary == null) return false;

            foreach (var key in keys)
            {
                if (_rowKeys.Contains(key))
                {
                    Console.WriteLine("already contianed");
                    return false;
                }
            }
            var newRows = new List<IList<object>>();

            foreach (var key in keys)
            {
                var listValues = new List<object>();
                foreach (var header in _headers)
                {
                    listValues.Add(dictionary[key][header]);
                }
                newRows.Add(listValues);
            }

            var request = GetSheetService().Spreadsheets.Values.Append(new ValueRange()
            {
                Values = newRows,
                MajorDimension = "ROWS"
            }, _spreadSheetId, _name);

            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;

            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;

            var response = await request.ExecuteAsync();

            Console.WriteLine(response.Updates.UpdatedRows);

            return true;
        }

        public async Task<bool> UpdateUICD(string key, Dictionary<string, string> dict)
        {
            var nonEmptyCells = await ReadFromSheet(SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE,
                SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER);

            var dat = nonEmptyCells;
            var rows = nonEmptyCells.Skip(1);
            var matchingRow = rows.FirstOrDefault(row => row[0].ToString() == key);

            if (matchingRow == null)
            {
                Console.WriteLine("no matching row");
                return false;
            }

            var rowToUpdate = new List<object>();

            foreach (var header in _headers)
            {
                rowToUpdate.Add(dict[header]);
            }

            var rowIndex = nonEmptyCells.IndexOf(matchingRow) + 1;

            var updateRange = $"{_name}!A{rowIndex}:Z{rowIndex}";
            var updateRequest = new ValueRange()
            {
                Range = updateRange,
                Values = new List<IList<object>> { rowToUpdate },
            };

            var updateResponse = GetSheetService().Spreadsheets.Values.Update(updateRequest, _spreadSheetId, updateRange);
            updateResponse.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            await updateResponse.ExecuteAsync();

            Console.WriteLine("Updated row successfully");

            return true;
        }

        public async Task<bool> UpdateUICD(string key, string columnKey, string value)
        {
            var nonEmptyCells = await ReadFromSheet(SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE,
                        SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.FORMATTEDSTRING);

            var rows = nonEmptyCells.Skip(1);
            var matchingRow = rows.FirstOrDefault(row => row[0].ToString() == key);


            if (matchingRow == null)
            {
                Console.WriteLine("no matching row");
                return false;
            }

            var columIndex = _headers.IndexOf(columnKey);


            var rowToUpdate = new List<object>();

            var rowIndex = nonEmptyCells.IndexOf(matchingRow) + 1;
            var updateRange = $"{_name}!{GetColumnCharachter(columIndex + 1)}{rowIndex}";
            var updateRequest = new ValueRange()
            {
                Range = updateRange,
                Values = new List<IList<object>> { new List<object> { value } },
            };

            var updateResponse = GetSheetService().Spreadsheets.Values.Update(updateRequest, _spreadSheetId, updateRange);
            updateResponse.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            await updateResponse.ExecuteAsync();

            Console.WriteLine("Updated row successfully");


            return true;
        }

        public bool deleteUICD(string key)
        {
            //null;
            return false;
        }

        public string GetRevision()
        {
            return "";
        }

        public Dictionary<string, Dictionary<string, string>> GetCurrentSheet() => _currentSheet;
        public List<string> GetRowKeys() => _rowKeys;
        public List<string> GetHeaders() => _headers;
        public string GetName() => _name;

        public Sheet(string spreadSheetId, string name)
        {
            _spreadSheetId = spreadSheetId;
            _name = name;
        }

        private SpreadSheet parent;

        /// <summary>
        /// IS SEEND IN BROWSER
        /// </summary>
        private string _spreadSheetId;

        /// <summary>
        /// NAME OF THE TAB
        /// </summary>
        private string _name;

        /// <summary>
        /// NAME OF THE TAB
        /// </summary>
        private string _range;
        /// <summary>
        /// THE FIRST COLOUMN OF EACH ROW
        /// </summary>
        private List<string> _rowKeys = new List<string>();

        /// <summary>
        /// HEADER COLUMN, THAT IS THE FIRST ROW IN SPREADSHEET
        /// </summary>
        private List<string> _headers;

        /// <summary>
        /// THE UIDCOLUMNBASEDMAPS (UICD)
        /// </summary>
        private Dictionary<string, Dictionary<string, string>> _currentSheet = new Dictionary<string, Dictionary<string, string>>();

        /// <summary>
        /// TWO DIMENSIOPNAL SPREADSHEET DATA ARAAY IN LISTO'LIST SHAPE
        /// </summary>
        private List<List<string>> _spreadSheetToListMapper = new List<List<string>>();

        private void SetDictionary()
        {
            if (_spreadSheetToListMapper == null || _spreadSheetToListMapper.Count == 0) return;

            _headers = _spreadSheetToListMapper[0];

            _rowKeys.Clear();

            for (int i = 1; i < _spreadSheetToListMapper.Count; i++)
            {
                if (_spreadSheetToListMapper[i].Count == 0) continue;

                string key = _spreadSheetToListMapper[i][0];

                Dictionary<string, string> valueDic = new Dictionary<string, string>();

                for (int j = 0; j < _headers.Count; j++)
                {
                    valueDic.Add(_headers[j], _spreadSheetToListMapper[i][j]);
                }

                _rowKeys.Add(key);

                _currentSheet[key] = valueDic;
            }

        }

        private async Task<List<List<object>>> ReadFromSheet(SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum uNFORMATTEDVALUE, SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum sERIALNUMBER)
        {
            SpreadsheetsResource.ValuesResource valuesResource;

            valuesResource = GetSheetService().Spreadsheets.Values;

            string ReadRange = $"{_name}!A:Z";

            var request = valuesResource.Get(_spreadSheetId, ReadRange);

            request.ValueRenderOption = uNFORMATTEDVALUE;
            request.DateTimeRenderOption = sERIALNUMBER;

            var response = await request.ExecuteAsync();

            var nonEmptyRows = response.Values.Where(row => row.Any());
            var nonEmptyCells = nonEmptyRows.Select(row => row.Where(cell => !string.IsNullOrEmpty(cell.ToString())));

            return nonEmptyCells.Select(row => row.ToList()).ToList();
        }

        private void SetList(List<List<object>> nonEmptyCells)
        {
            _spreadSheetToListMapper.Clear();

            foreach (var row in nonEmptyCells)
            {
                List<string> stringList = row.OfType<string>().ToList();
                _spreadSheetToListMapper.Add(stringList);
            }
        }

        private SheetsService GetSheetService()
        {
            using (var stream = new FileStream(SpreadSheet.GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(SpreadSheet.Scopes)
                };

                return new SheetsService(serviceInitializer);
            }
        }

        private string GetColumnCharachter(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
