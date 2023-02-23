using GoogleSheetApi;



/// <summary>
/// Uncomment this code to readSpreadsheet
/// </summary>

string spreadSheetID = string.Empty;

CommandsFileCreator.filePath = GetFilePath();
ChooseSpreadsheetToWorkOn();
SpreadSheet.Id = spreadSheetID;

SpreadSheet spreadSheet = new SpreadSheet();
var tabs = await spreadSheet.GetSUICD();
CommandsFileCreator.CreateCommandFile(tabs, spreadSheet.GetSheets());
Console.WriteLine("success");

//while (true)
//{
//    Thread.Sleep(2000);
//    var updatedtabsrows = await spreadSheet.GetDeltaSUICD();
//    CommandsFileCreator.CreateCommandFile(updatedtabsrows, spreadSheet.GetSheets());

//}

void ChooseSpreadsheetToWorkOn()
{
    Console.WriteLine("Please enter the number of the Team's spreadsheet you want to work on : ");
    Console.WriteLine("=========================================================================");
    Console.WriteLine("1 - Aristocraten Amsterdam");
    Console.WriteLine("2 - Breda Beesten");
    Console.WriteLine("3 - Eindhoven Vanguards");
    Console.WriteLine("4 - Groningse Geleerden");
    Console.WriteLine("5 - De Haagse Helden");
    Console.WriteLine("6 - KC Haarlemmer Wieken");
    Console.WriteLine("7 - Haven Strijders Rotterdam ");
    Console.WriteLine("8 - Utrecht Ballerinas");


    var index = Console.ReadLine();

    switch (index)
    {
        case "1":
            spreadSheetID = "1jJKfSwuOpKbNbxvYh4ZPKj999wsVZfMq2pUmxlBC4r8";
            SpreadSheet.Name = "Aristocraten Amsterdam";
            break;
        case "2":
            spreadSheetID = "1v3JqKUOW2BBlnxOavqlT1-JgjYTgv_uCSuS0OAcXAVA";
            SpreadSheet.Name = "Breda Beesten";
            break;
        case "3":
            spreadSheetID = "1P2Ih3lBlpPN5g3B7ZgtQmb6tImI-4X1HPR548BeJJ8A";
            SpreadSheet.Name = "Eindhoven Vanguards";
            break;
        case "4":
            spreadSheetID = "1kxklPGTKD27MR_Dw1DBYvq34qPI9Ma5NAzRQaEa4fUk";
            SpreadSheet.Name = "Groningse Geleerden";
            break;
        case "5":
            spreadSheetID = "1E3JcwTF_qDCtWsGmWNJZlEK0VhxdMwL5vS49vL1akEI";
            SpreadSheet.Name = "De Haagse Helden";
            break;
        case "6":
            spreadSheetID = "1PWxDXOeR1e5KC6Ek1Jmuq4Ixg_Ol9tJt0TC7NVJqJNo";
            SpreadSheet.Name = "KC Haarlemmer Wieken";
            break;
        case "7":
            spreadSheetID = "1P24GghaYWL9RLxtjmhAygJMLkLStsG6Fqfztfs54viY";
            SpreadSheet.Name = "Haven Strijders Rotterdam ";
            break;
        case "8":
            spreadSheetID = "1FRct0ewdoPkgVB_5_KX77kZ3P18rEmXOAsmwFVyL0ZQ";
            SpreadSheet.Name = "Utrecht Ballerinas";
            break;
        default:
            break;
    }
}

string GetFilePath()
{
    string projectFolder = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory()));
    projectFolder = projectFolder.Substring(0, projectFolder.LastIndexOf("\\"));
    return projectFolder + "/CommandFiles";
}


/// <summary>
/// Uncomment this code to Append rows in spreadsheet also remove row with ID of 4 in spreadsheet to test it or put 5 instead of 4 or any other higher value
/// </summary>


//Sheet main = new Sheet("1jJKfSwuOpKbNbxvYh4ZPKj999wsVZfMq2pUmxlBC4r8", "Main");
//await main.GetUICD();
//Dictionary<string, Dictionary<string, string>> appenvalues = new Dictionary<string, Dictionary<string, string>>();
//Dictionary<string, string> valudedic = new Dictionary<string, string>();
//foreach (var header in main.GetHeaders())
//{
//    if (header == "ID")
//    {
//        valudedic[header] = "4";
//        continue;
//    }
//    Random random = new Random();
//    valudedic[header] = random.Next().ToString();
//}

//appenvalues["4"] = valudedic;

//await main.ApendUICD(appenvalues, new List<string>() { "4" });

/// <summary>
/// Uncomment this code to update the whole row of existing data with new values
/// </summary>

//Sheet Main = new Sheet("1jJKfSwuOpKbNbxvYh4ZPKj999wsVZfMq2pUmxlBC4r8", "Main");
//await Main.GetUICD();

//Dictionary<string, string> dic = new Dictionary<string, string>();
//string key = "4";
//var headers = Main.GetHeaders();
//foreach (var header in headers)
//{
//    if (header == "ID") dic[header] = key;

//    else
//    {
//        dic[header] = "randomingg";

//    }

//}
//await Main.UpdateUICD(key, dic);


/// <summary>
/// Uncomment this code to update one specific cell using ro number and column name
/// </summary>

//Sheet main = new Sheet("1jJKfSwuOpKbNbxvYh4ZPKj999wsVZfMq2pUmxlBC4r8", "Main");
//await main.GetUICD();



//string key = "4";
//string columname = "Value1";
//string vallue = "Noah";
//await main.UpdateUICD(key, columname, vallue);




#region FileSystem

//var files =GoogleSheetsV2.ListFiles();
//foreach (var item in files.Files)
//{
//    Console.WriteLine(item.Name);

//}
//Directory.CreateDirectory("LOGOS");
//GoogleSheetsV2.syncAllFileToDirecotry("LOGOS");


#endregion