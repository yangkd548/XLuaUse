using System.Collections.Generic;
using System.IO;
using System.Text;
using LitJson;
using OfficeOpenXml;
using UnityEditor;
using UnityEngine;
public class Xls2Lua :  EditorWindow{
	private static Dictionary<string,ExcelWorksheet> dict;
	private static string dstPath;
	private static string filePath;
	private static string fileName;
	private static string upperFileName;
	private static string luaHead;
	private static string luaMiddle1;
	private static string luaMiddle2;
	private static string luaMiddle3;
	private static string luaFoot;
	[MenuItem("Tools/一键导表", false, 1)]
	public static void StartConvert()
    {
		string srcPath = Application.dataPath + "/../xls";
		DirectoryInfo dir = new DirectoryInfo(srcPath);
		dstPath = Application.dataPath + "/Lua/data";

		List<FileInfo> fileList = new List<FileInfo>();
		List<string> importStrList = new List<string>();

		foreach (FileInfo f in dir.GetFiles())
		{
			string fileNameLower = f.Name.ToLower();
            if (f.Extension == ".xlsx" && fileNameLower.Substring(0,1) != "~" && fileNameLower.Substring(fileNameLower.Length -7,2) == "_c")
            // if (f.Name == "skill_C.xlsx")
            {
				fileList.Add(f);
			}
		}
		for (int i = 0; i < fileList.Count; i++)
		{
			var f = fileList[i];
			EditorUtility.DisplayProgressBar("导表("+i + "/" + fileList.Count + ")", f.Name, i / (float)fileList.Count);
			try
			{
				ConvertXls2Lua(f.DirectoryName,f.Name);
				importStrList.Add(upperFileName + " = require(\"data." + upperFileName + "\").new()");
			}
			catch (System.Exception)
			{
				EditorUtility.ClearProgressBar();
				EditorUtility.DisplayDialog("导表出错",f.Name + "格式有误，检查一下吧！","OK");
				Debug.LogError(f.Name + "格式有误，检查一下吧！");
				throw;
			}
		}
		if (fileList.Count > 1)
		{
			File.WriteAllBytes(dstPath + "/ImportDataTable.lua", Encoding.UTF8.GetBytes(string.Join("\n",importStrList.ToArray())));
		}
		EditorUtility.ClearProgressBar();
		AssetDatabase.Refresh();
		Debug.Log("导表完毕");
    }
	public static void ConvertXls2Lua(string path,string name){
		// Debug.Log("---> "+name);
		filePath = path;
		fileName = name.Substring(0,name.Length - 5);
		upperFileName = fileName.Substring(0,1).ToUpper() + fileName.Substring(1,fileName.Length - 1);
		InitLuaString();
		ReadXls(path,name);
	}
	private static void ReadXls(string path,string name)
    {
		dict = new Dictionary<string,ExcelWorksheet>();
		path += "/"+name;
		using(ExcelPackage package = new ExcelPackage(new FileStream(path, FileMode.Open,FileAccess.Read)))
		{
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
			ConvertData();
		}
    }

	private static void ConvertData(){
		string luaStr = "";
		Dictionary<string,int> sheetTitleDict = new Dictionary<string,int>();
		ExcelWorksheet sheet = dict["Sheet1"];
		ExcelWorksheet def = dict["def"];
		List<string> totalLineList = new List<string>();
		// Sheet1表字段索引
		List<string> luaKeyList = new List<string>();
		for ( int j = def.Dimension.Start.Column, k = def.Dimension.End.Column; j <= k; j++ )
		{
			string strDef = GetCellValue(def,1, j);
			if ( strDef != "nil" )
			{
				for ( int m = sheet.Dimension.Start.Column, n = sheet.Dimension.End.Column; m <= n; m++ )
				{
					string strSht = GetCellValue(sheet,1,m);
					if (strSht != "nil" && strDef == strSht)
					{
						sheetTitleDict[strDef] = m;
						luaKeyList.Add("        [\"" + GetCellValue(def,2, j) + "\"]=" + sheetTitleDict.Count);
						break;
					}
				}
			}
		}
		// def字典
		Dictionary<string,DefData> defDict = new Dictionary<string,DefData>();
		DefData defData;
		for ( int i = def.Dimension.Start.Column, j = def.Dimension.End.Column; i <= j; i++ )
		{
			string strDef = GetCellValue(def,1, i);
			if (strDef != "nil")
			{
				defData = new DefData();
				defData.inited = true;
				defData.name = strDef;
				defData.key = GetCellValue(def,2, i);
				defData.valueType = GetCellValue(def,3, i).ToLower();
				defDict.Add(defData.name,defData);
			}
		}
		string luaTransStr = "";
		foreach (KeyValuePair<string, DefData> item in defDict)
		{
			if (item.Value.valueType == "string$")
			{
				luaTransStr += item.Value.key + "=1,";
			}
		}
		if (luaTransStr != "")
		{
			luaTransStr = luaTransStr.Substring(0,luaTransStr.Length - 1);
		}
		// 遍历行
		for ( int i = 2, j = sheet.Dimension.End.Row; i <= j; i++ )
		{
			if (GetCellValue(sheet,i,1) == "nil")
			{
				break;
			}
			// 遍历列
			List<string> lineStrList = new List<string>();
			foreach (KeyValuePair<string, DefData> item in defDict)
			{
				if (item.Value.inited)
				{
					try
					{
						lineStrList.Add(ConvertAllType(item.Value,GetCellValue(sheet,i,sheetTitleDict[item.Value.name])));
					}
					catch (System.Exception)
					{
						Debug.Log("导出\""+fileName+": 第"+i+"行,"+item.Value.name+"\"出错!");
						throw;
					}
				}
			}
			string temp = string.Join(",",lineStrList.ToArray());
			temp = temp.Replace("\n","\\n");
			temp = temp.Replace("\r","\\r");
			string line = "";
            if (defDict["id"].valueType == "number")
            {
                line = "        [" + GetCellValue(sheet,i,sheetTitleDict["id"]) + "]={";
            }
            else
            {
                line = "        [\"" + GetCellValue(sheet,i,sheetTitleDict["id"]) + "\"]={";
            }
			line += temp;
			line += "}";
			line = line.Replace(",}","}");
			totalLineList.Add(line);
		}
		luaStr = string.Join(",\n",totalLineList.ToArray());
		string luaIsArrStr;
		DefData defTemp = defDict["id"];
		if (defTemp.valueType == "number")
		{
			luaIsArrStr = "true";
		}
		else
		{
			luaIsArrStr = "false";
		}
		CreateLua(luaHead + string.Join(",\n",luaKeyList.ToArray()) + luaMiddle1 + luaStr + luaMiddle2 + luaTransStr + luaMiddle3 + luaIsArrStr + luaFoot);
	}
	private static string ConvertAllType(DefData defData,string value)
	{
		string result = "";
		// Debug.Log(value);
		// if (defData.key == "initialskill")
		// {
		// 	int a=1;
		// }
		if (defData.valueType == "string" || defData.valueType == "string$")
		{
			if (value != "nil")
			{
				if (value.IndexOf("\"") >= 0)
				{
					value = value.Replace("\"","\\"+"\"");
				}
				result += "\"" + value + "\"" ;
			}
			else
			{
				result += "\"\"" ;
			}
		}
		else if (defData.valueType == "number")
		{
			if (value != "nil" && value != "")
			{
				result += value;
			}
			else
			{
				result += "0";
			}
		}
		else if(defData.valueType == "bool")
		{
			if (value == "1")
			{
				result += "true";
			}
			else
			{
				result += "false";
			}
		}
		else if(defData.valueType == "json")
		{
			List<string> strList = new List<string>();
			if (value == "nil" || (value.Substring(0,1) != "[" && value.Substring(0,1) != "{"))
			{
				result += "{}";
			}
			else
			{
				JsonData jd = JsonMapper.ToObject(value);
				ParseJson(jd,ref strList);
				strList.RemoveAt(strList.Count - 1);
				result += string.Join("",strList.ToArray());
			}
		}
		else
		{
			Debug.Log("数据类型未知 "+defData.key + ":" + defData.valueType);
			throw new System.Exception("数据类型未知");
		}
		return result;
	}
	public static string GetCellValue(ExcelWorksheet sheet,int i,int j)
	{
		object o = sheet.GetValue(i,j);
		if (o == null)
		{
			return "nil";
		}
		else
		{
			return o.ToString();
		}
	}
	public static void ParseJson(JsonData jsonData,ref List<string> strList,string key = ""){
		if (key != "")
		{
			strList.Add(key + "=");
			if (jsonData.GetJsonType() == JsonType.Array)
			{
				strList.Add("{");
				foreach (JsonData p in jsonData){
					ParseJson(p,ref strList);
				}
				strList.Add("}");
			}
			else if (jsonData.GetJsonType() == JsonType.Object)
			{
				strList.Add("{");
				foreach (KeyValuePair<string, JsonData> item in jsonData)
				{
					ParseJson(item.Value,ref strList,item.Key);
				}
				strList.Add("}");
			}
			else if (jsonData.GetJsonType() == JsonType.String)
			{
				strList.Add("\"");
				strList.Add(jsonData.ToString());
				strList.Add("\"");
			}
			else if (jsonData.GetJsonType() == JsonType.Int)
			{
				strList.Add(jsonData.ToString());
			}
			else if (jsonData.GetJsonType() == JsonType.Long)
			{
				strList.Add(jsonData.ToString());
			}
			else if (jsonData.GetJsonType() == JsonType.Double)
			{
				strList.Add(jsonData.ToString());
			}
			else if (jsonData.GetJsonType() == JsonType.Boolean)
			{
				if (jsonData.ToString() == "True")
				{
					strList.Add("true");
				}
				else
				{
					strList.Add("false");
				}
			}
		}
		else
		{
			if (jsonData.GetJsonType() == JsonType.Array)
			{
				strList.Add("{");
				foreach (JsonData p in jsonData){
					ParseJson(p,ref strList);
				}
				strList.Add("}");
			}
			else if (jsonData.GetJsonType() == JsonType.Object)
			{
				strList.Add("{");
				foreach (KeyValuePair<string, JsonData> item in jsonData)
				{
					ParseJson(item.Value,ref strList,item.Key);
				}
				strList.Add("}");
			}
			else if (jsonData.GetJsonType() == JsonType.String)
			{
				strList.Add("\"");
				strList.Add(jsonData.ToString());
				strList.Add("\"");
			}
			else if (jsonData.GetJsonType() == JsonType.Int)
			{
				strList.Add(jsonData.ToString());
			}
			else if (jsonData.GetJsonType() == JsonType.Long)
			{
				strList.Add(jsonData.ToString());
			}
			else if (jsonData.GetJsonType() == JsonType.Double)
			{
				strList.Add(jsonData.ToString());
			}
			else if (jsonData.GetJsonType() == JsonType.Boolean)
			{
				if (jsonData.ToString() == "True")
				{
					strList.Add("true");
				}
				else
				{
					strList.Add("false");
				}
			}
		}
		strList.Add(",");
	}
	private static void CreateLua(string luaStr){
		File.WriteAllBytes(dstPath + "/" + upperFileName + ".lua", Encoding.UTF8.GetBytes(luaStr));
	}
	private static void InitLuaString(){
		luaHead = "local " + upperFileName + "=class(\"" + upperFileName + "\")\n" +
		"local datas = nil\n" +
		"local keys = nil\n" +
		"local isArray = nil\n" +
		"function " + upperFileName + ":ctor()\n" +
		"	keys={\n";
		luaMiddle1 = "\n	}\n" +
		"	datas = {\n";
		luaMiddle2 = "\n	}\n" +
		"	self._transKeyDict = {";
		luaMiddle3 = "}\n" +
		"	isArray = ";
		luaFoot = "\nend\n" +
		"local function getData(id)\n" +
        "    if not id then\n" +
        "        return\n" +
        "    end\n" +
        "    local key,obj\n" +
        "    if isArray then\n" +
        "        key = tonumber(id)\n" +
        "    else\n" +
        "        key = tostring(id)\n" +
        "    end\n" +
        "    obj = datas[key]\n" +
        "    if not obj then\n" +
        "        if isArray then\n" +
        "            key = -1\n" +
        "        else\n" +
        "            key = \"-1\"\n" +
        "        end\n" +
        "        obj = datas[key]\n" +
        "    end\n" +
        (upperFileName == "Word_C"? "":
        "    if obj == nil and not gConfig.release then\n" +
        "        error(\"" + upperFileName + ".id = \"..id..\"不存在，请 【修改-数据表】、【导表】、【改代码】（尽量不要改id）\")\n" +
        "    end\n") +
        "    return obj\n" +
		"end\n" +
		"function " + upperFileName + ":createData(data_)\n" +
		"    local data = {}\n" +
		"    for k,v in pairs(keys) do\n" +
		"        data[k] = data_[v]\n" +
		"    end\n" +
		"    return data\n" +
		"end\n" +
		"function " + upperFileName + ":getRecord(id)\n" +
		"	local data = getData(id)\n" +
		"	if data == nil then\n" +
		"		return\n" +
		"	end\n" +
		"	return self:createData(data)\n" +
		"end\n" +
		"function " + upperFileName + ":getValue(id, key)\n" +
		"    local record = getData(id)\n" +
		"    if record then\n" +
		"        if self._transKeyDict[key] == 1 then\n" +
		"            return gTransMgr:getString(record[keys[key]])\n" +
		"        else\n" +
		"            return record[keys[key]]\n" +
		"        end\n" +
		"    end\n" +
		"end\n" +
		"function " + upperFileName + ":getIdx()\n" +
		"	local arr = {}\n" +
		"	for k,v in pairs(datas) do\n" +
		"		table.insert(arr,k)\n" +
		"	end\n" +
		"	return arr\n" +
		"end\n" +
		"return " + upperFileName;
	}
}
[System.Serializable]
public class DefData {
	public bool inited = false;
	public string name;
	public string key;
	public string valueType;
}
[System.Serializable]
public class SheetData {
	public List<List<string>> datas = new List<List<string>>();
}