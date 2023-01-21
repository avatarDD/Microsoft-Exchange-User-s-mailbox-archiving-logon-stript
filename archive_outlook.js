//======================================================USER'S LOGON SCRIPT FOR ARCHIVING EXCHANGE MAILBOX==================
var monthCount = 6;                        //за сколько месяцев оставлять почту
var defaultPathToPST = "d:\\АрхивПочты";    //путь для архива почты по-умолчанию (если нет диска D, тогда файл .PST будет создан в папке "мои документы" на диске С)
//========================================================================


var olMail = 43;                            //Объект типа MailItem.

var dateLine = new Date(new Date().getTime() - (monthCount * 30 * 24 * 60 * 60 * 1000));
var dateLineTxt = convertDate(dateLine);

var fso = WScript.CreateObject("Scripting.FileSystemObject");
var WshShell = WScript.CreateObject("WScript.Shell");
var objADSystemInfo = WScript.CreateObject("ADSystemInfo");
var objOL = null;
try
{
    objOL = WScript.CreateObject("Outlook.Application");
} catch (err)
{
    //WScript.Echo(err.description); //вывести сообщение об ошибке
    WScript.Quit(1);
}
//=========для отладки===============
//printAllDir(objOL.Session.Folders);
//var curUsrMail = "user@company.com";       //название профиля outlook (обычно совпадает с адресом e-mail)
//===================================

var curUsrMail = GetObject("LDAP://" + objADSystemInfo.UserName).mail;

//delOSTfile();
var archDir = FindOrCreatePST(objOL);
while (archDir == null)
{   //---после создания нового .pst файла надо перезапустить outlook
    try
    {
        objOL.Quit();
    } catch (err) { }
    objOL = WScript.CreateObject("Outlook.Application");
    archDir = FindOrCreatePST(objOL);
}

for (var i = 1; i < objOL.Session.Folders.Count+1; i++)
{
    if (objOL.Session.Folders(i).Name.toLowerCase() == curUsrMail.toLowerCase())
    {//----нашли нужный нам профиль---
        createkSameDirsInArchAndMoveMSGs(objOL.Session.Folders(i), archDir);    //зеркалируем структуру папок и перемещаем старые письма в .pst
        break;
    }
}

//objOL.Quit(); //закрыть outlook по готовности

//=========================================================================================================================================
function createkSameDirsInArchAndMoveMSGs(d, ad)
{
    for (var i = 1; i < d.Folders.Count + 1; i++)
    {
        if (d.Folders(i).Name.toLowerCase() == "календарь" ||
            d.Folders(i).Name.toLowerCase() == "удаленные" ||
            d.Folders(i).Name.toLowerCase() == "черновики" ||
            d.Folders(i).Name.toLowerCase() == "ошибки синхронизации")
            continue;

        if (d.Folders(i).Items.Count > 0)
        {
            try
            {
                ad.folders.Add(d.Folders(i).Name);
                //WScript.Echo("создана новая папка '" + d.Folders(i).Name + "'");

            } catch (err)
            {
                //WScript.Echo("папка '" + d.Folders(i).Name + "' уже есть");
            }
        }
        //*****************************************************************        
        var MsgList = d.Folders(i).Items();

        if (MsgList.Count > 0)
        {
            if (MsgList(1).Class == olMail)
            {//----это письма, а не календари и проч.
                if (d.Folders(i).Name.toLowerCase() == "отправленные")
                {//-------отбор отправленных по дате отправки-----
                    var outMsgsList = RestrictMessages(MsgList, "[SentOn]", dateLineTxt);
                    MoveMessages(outMsgsList, d.Folders(i), ad.Folders(d.Folders(i).Name),true);
                } else
                {//-------отбор входящих по дате получения-------
                    var outMsgsList = RestrictMessages(MsgList, "[ReceivedTime]", dateLineTxt);
                    MoveMessages(outMsgsList, d.Folders(i), ad.Folders(d.Folders(i).Name),false);
                }
            }
            createkSameDirsInArchAndMoveMSGs(d.Folders(i), ad.folders(d.Folders(i).Name));
        }
        //*****************************************************************
    }    
}

function RestrictMessages(msgsLst,fltr,dtln)
{    
    var strFilter = fltr + " < '" + dtln + "'";
    var MsgListFiltered = msgsLst.Restrict(strFilter);
    MsgListFiltered.Sort(fltr);
    return MsgListFiltered;
}

function MoveMessages(msgsLst,src,dst,isSentBox)
{
    for (var k = 1; k < msgsLst.Count + 1; k++)
    {
        if (msgsLst(k).Class == olMail)
        {
            try
            {
                msgsLst(k).Move(dst);
            } catch (err)
            {//-----при большом количестве писем может сработать защита сервера на количество одновременно открытых писем----
                //WScript.Echo(err.description); //вывести сообщение об ошибке
                WScript.Quit(1);
            }
            //=======================для отладки==================================
            /*
                var text = "";
                if (isSentBox)
                {
                    text = "Перемещено письмо: '" + msgsLst(k).Subject + "' от " + convertDate(msgsLst(k).SentOn) + "\r\nиз '" + src.FolderPath + "' в '" + dst.FolderPath + "'";

                } else
                {
                    text = "Перемещено письмо: '" + msgsLst(k).Subject + "' от " + convertDate(msgsLst(k).ReceivedTime) + "\r\nиз '" + src.FolderPath + "' в '" + dst.FolderPath + "'";
                }
                WScript.Echo(text);
            */
            //====================================================================
        }
    }
}

function FindOrCreatePST(ol)
{
    var result = null;

    try
    {

        for (var i = 1; i < ol.Session.Stores.Count + 1; i++)
        {
            if (/archive[^\\]*.pst$/gi.test(ol.Session.Stores.Item(i).FilePath))
            {//---нашли archive.pst----
                result = ol.Session.folders(ol.Session.Stores.Item(i).DisplayName);
                var obf = result.Store.GetRootFolder();
                obf.Name = "Архивы"; //переименовали профиль с pst  
                return result;
            }
        }
        if (result == null)
        {  //----не нашли, создаём новый archive.pst    
            var archPath = "";
            if (fso.FolderExists("d:"))
            {
                archPath += defaultPathToPST + "\\" + curUsrMail + "\\Archive.pst";
            }else
            {
                if (getOS() == "win7")
                {
                    archPath += WshShell.ExpandEnvironmentStrings("%UserProfile%") + "\\Documents\\Файлы Outlook" + "\\Archive.pst";
                }else
                {
                    archPath += WshShell.ExpandEnvironmentStrings("%UserProfile%") + "\\Мои документы\\Файлы Outlook" + "\\Archive.pst";
                }
            }
            ol.Session.AddStoreEx(archPath, 2);  //OlStoreType 1-default, 2-unicode, 3-ansi
            FindOrCreatePST(ol);
        }
    } catch (err)
    {
        return result;
    }
}

function getOS()
{
    var OS = WshShell.RegRead("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProductName");

    if (/windows (7|8|8\.1|10|11)/gi.test(OS))
    {
        OS = "win7";
    }
    else
    {
        if (/windows xp/gi.test(OS))
        {
            OS = "winxp";
        }
    }
    return OS;
}

function convertDate(dt)
{
    var d = new Date(dt);
    var day = d.getDate() < 10 ? "0" + d.getDate() : d.getDate();
    var month = (d.getMonth() + 1 < 10) ? "0" + (d.getMonth() + 1) : (d.getMonth() + 1);
    var year = d.getYear();
    return day + "." + month + "." + year;
}

function delOSTfile()
{
    var OSTPath = "";
    if (getOS() == "win7")
	{
        OSTPath += WshShell.ExpandEnvironmentStrings("%UserProfile%") + "\\AppData\\Local\\Microsoft\\Outlook";
    } else {
        OSTPath += WshShell.ExpandEnvironmentStrings("%UserProfile%") + "\\Local Settings\\Application Data\\Microsoft\\Outlook";
    }
    if (fso.FolderExists(OSTPath))
    {
        var d = fso.GetFolder(OSTPath);
        var files = d.Files;

        for (var f = new Enumerator(files); !f.atEnd(); f.moveNext())
        {
            if (/\.ost$/.test(f.item().Path))
            {
                killOutlookProcess();
                fso.DeleteFile(f.item().Path, true);
            }
        }
    }
}

function killOutlookProcess()
{
    WshShell.Run("taskkill /f /im outlook.exe", 0, true);
}

function printAllDir(dir)
{    //---вывести все папки и подпапки
    for (var i = 1; i < dir.Count + 1; i++)
    {
        WScript.Echo(dir(i).FolderPath);
        printAllDir(dir(i).Folders);
    }
}
