using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;

public class Xlsx{
	private Dictionary<string, ExcelWorksheet> dict;
	private Dictionary<string, int> sheetTitleDict;
	public ExcelWorksheet sheet;
	private ExcelWorksheet def;
	private Dictionary<string, XlsDefData> defDict;

	public Xlsx(string path)
    {
		dict = new Dictionary<string, ExcelWorksheet>();
		ExcelPackage package = new ExcelPackage(new FileStream(path, FileMode.Open,FileAccess.Read));
		for ( int i = 1; i <= package.Workbook.Worksheets.Count; ++i )
		{
			ExcelWorksheet sheet = package.Workbook.Worksheets[i];
			if (sheet.Name == "Sheet1" || sheet.Name == "def")
			{
				if (!dict.ContainsKey(sheet.Name))
				{
					dict[sheet.Name] = sheet;
				}
			}
		}
		this.GetDef();
    }
    public void GetDef()
    {
		sheetTitleDict = new Dictionary<string, int>();
		sheet = dict["Sheet1"];
		def = dict["def"];
		List<string> totalLineList = new List<string>();
		// Sheet1表字段索引
		List<string> luaKeyList = new List<string>();
		for (int j = def.Dimension.Start.Column, k = def.Dimension.End.Column; j <= k; j++)
		{
			string strDef = GetCellValue(def, 1, j);
			if (strDef != "nil")
			{
				for (int m = sheet.Dimension.Start.Column, n = sheet.Dimension.End.Column; m <= n; m++)
				{
					string strSht = GetCellValue(sheet, 1, m);
					if (strSht != "nil" && strDef == strSht)
					{
						sheetTitleDict[strDef] = m;
						luaKeyList.Add("        [\"" + GetCellValue(def, 2, j) + "\"]=" + sheetTitleDict.Count);
						break;
					}
				}
			}
		}
		// def字典
		defDict = new Dictionary<string, XlsDefData>();
		XlsDefData defData;
		for (int i = def.Dimension.Start.Column, j = def.Dimension.End.Column; i <= j; i++)
		{
			string strDef = GetCellValue(def, 1, i);
			if (strDef != "nil")
			{
				defData = new XlsDefData
				{
					inited = true,
					name = strDef,
					key = GetCellValue(def, 2, i),
					valueType = GetCellValue(def, 3, i).ToLower(),
					col = i
                };
                defDict.Add(defData.name, defData);
			}
		}
	}
	public string GetCellValue(ExcelWorksheet sheet, int i, int j)
	{
		object o = sheet.GetValue(i, j);
		if (o == null)
		{
			return "nil";
		}
		else
		{
			return o.ToString();
		}
	}
    public string GetValue(string name,int row)
    {
		string result = GetCellValue(sheet, row, sheetTitleDict[defDict[name].name]);
		return result;
    }
    public bool HasId(string id)
    {
		bool result = false;
		for (int row = 2, col = sheet.Dimension.End.Row; row <= col; row++)
        {
            if (id == GetCellValue(sheet,row, sheetTitleDict[defDict["id"].name]))
            {
				result = true;
				break;
			}
        }
		return result;
    }
}
public class XlsDefData
{
	public bool inited = false;
	public string name;
	public string key;
	public string valueType;
	public int col;
}