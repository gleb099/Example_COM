using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Threading;

using CADLibKernel;
using nanoCAD;

using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ComWrapper1C
{
    //Интерфейс для COM-компоненты
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsDual), Guid("3952896B-E770-4CCE-B25B-F5C979F9CADE")]
    public interface IComClass
    {
        [DispId(1)]
        bool createConnection(string connectionInfo);
        [DispId(2)]
        string addFolders(string objectName, string parentObjectGUID = null);
        [DispId(3)]
        string addFolder(string objectName, string parentObjectGUID = null);
        [DispId(4)]
        string addDocCard(string objectName, string parentObjectGUID = null);
        [DispId(5)]
        void objectRestoration(string objectGUID);
        [DispId(6)]
        string addDocPart(string objectName, string parentObjectGUID);
        [DispId(7)]
        bool renameObject(string NewName, string objectGUID);
        [DispId(8)]
        void deleteObject(string objectGUID);
        [DispId(9)]
        void renameFile(string newName, string fileGUID);
        [DispId(10)]
        void updateFileContent(string fileGUID, string filePath);
        [DispId(11)]
        void setFileParameter(string fileGUID, string parameterName, object value);
        [DispId(12)]
        void setObjectParameter(string objectGUID, string parameterName, object value);
        [DispId(13)]
        void getFile(string fileGUID, string path);
        [DispId(14)]
        void openCADLibFileInModelStudio(string fileGUID);
        [DispId(15)]
        void openPathFileInModelStudio(string filePath);
        [DispId(17)]
        string addFile(string parentGUID, string filePath, string fileCategoryCaption);
        [DispId(18)]
        void deleteFile(string fileGUID);
        [DispId(19)]
        string createProjectInCADLib(string dbName);
        [DispId(20)]
        void openProjectInCADLib(string connectionInfo);
        [DispId(22)]
        void addToArchive(string objectGUID);
        [DispId(23)]
        void cadlibLink(string link, string connectionInfo);
    }

    //Основной класс COM-компоненты
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), Guid("7FC783E3-922D-499E-8A40-FFE9F6D62E29"), ProgId("ComWrapper1C.ComClass")]
    public class ComClass : IComClass
    {
        //Поля
        private CADLibraryBase library = new CADLibraryBase();
        //private CADLibraryBase library = null;
        private nanoCAD.Application ncApp = null;
        //Пустой конструктор(обязательно требуется) для регистрации в реестре
        public ComClass()
        {
        }

        #region Основные функции COM-компоненты

        /// <summary>
        /// Открытие в CADLib нужного проекта и отображдение требуемых элементов (только для 3д объектов). Иерархию не открывает, только отображает.
        /// </summary>
        /// <param name="link">Гиперссылка из CADLib МиА</param>
        public void cadlibLink(string link, string connectionInfo)
        {
            Dictionary<string, string> info = getCoonectInfo(connectionInfo);
            //string[] counts = link.Split(Convert.ToChar("[clLink>"));
            System.Text.StringBuilder temp = new System.Text.StringBuilder();
            temp.Append(link);
            temp.Append("[clLink<");
            string test = String.Join("", temp);
            int objCount = 0;
            //MessageBox.Show(test);
            for (int i = 0; i < test.Count(); i++)
            {
                if (test[i] == '[')
                {
                    if (test[i + 1] == 'c' && test[i + 2] == 'l' && test[i + 3] == 'L' && test[i + 4] == 'i' && test[i + 5] == 'n' && test[i + 6] == 'k' && test[i + 7] == '>')
                    {
                        objCount += 1;
                    }
                }
            }
            //string[] counts = link.Split('[clLink>');
            string[] links = link.Split('/');
            //int objCount = counts.Length;
            if (objCount >= 1)
            {
                Array.Clear(links, 0, links.Length);
                System.Text.StringBuilder res = new System.Text.StringBuilder();
                for (int i = 0; i < link.Length; i++)
                {
                    if (link[i] == '{')
                    {
                        while (link[i] != '}')
                        {
                            res.Append(link[i]);
                            i++;
                        }
                        res.Append('}');
                        if (objCount > 1) res.Append('|');
                        objCount--;
                    }
                }
                if (info["db_login"] == "null")
                {
                    Process cadlib = Process.Start($"clp:server:{info["server"]};db:{info["db"]};platform:mssql;auth:windows;show:selected;objects:{String.Join("", res)};");
                }
                else
                {
                    Process cadlib = Process.Start($"clp:server:{info["server"]};db:{info["db"]};platform:mssql;auth:dbms;show:selected;objects:{String.Join("", res)};");
                }
            }
        }

        /// <summary>
        /// Запуск NanoCad Model Studio CS и открытие в нем требуемого файла из БД CADLib
        /// </summary>
        /// <param name="fileGUID">GUID файла, который требуется открыть в NanoCad Model Studio CS</param>
        public void openCADLibFileInModelStudio(string fileGUID)
        {
            int idFile = library.GetFileIdByGUID(fileGUID);
            string fileName = library.GetFileInfoById(idFile).mFileName;
            library.DownloadFile((int)idFile, $"C:/Windows/Temp/{fileName}"); //temp

            try
            {
                ncApp = (nanoCAD.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("nanoCADx64.Application.20.0");
            }
            catch
            {
                if (ncApp == null)
                {
                    MessageBox.Show("Запуск может продолжиться около минуты");
                    Process ms = new Process();
                    ms.StartInfo.FileName = "C:/Program Files/Nanosoft/nanoCAD x64 Plus 20.3/nCad.exe";
                    ms.StartInfo.Arguments = $"-g \"C:/Program Files/CSoft/Model Studio CS/NANOPIPING/Pipe20.package\" -r Pipe";
                    ms.Start();
                    Thread.Sleep(60000);
                }
            }
            ncApp = (nanoCAD.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("nanoCADx64.Application.20.0");
            //ncApp.LoadModule("C:/Program Files/CSoft/Model Studio CS/NANOPIPING/Pipe20.package");
            ncApp.Documents.Open($"C:/Windows/Temp/{fileName}");
        }

        /// <summary>
        /// Запуск NanoCad Model Studio CS и открытие в нем требуемого файла из БД CADLib
        /// </summary>
        /// <param name="filePath">Путь к файлу, который требуется открыть в NanoCad Model Studio CS</param>
        public void openPathFileInModelStudio(string filePath)
        {
            try
            {
                ncApp = (nanoCAD.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("nanoCADx64.Application.20.0");
            }
            catch
            {
                if (ncApp == null)
                {
                    MessageBox.Show("Запуск может продолжиться около минуты");
                    Process ms = new Process();
                    ms.StartInfo.FileName = "C:/Program Files/Nanosoft/nanoCAD x64 Plus 20.3/nCad.exe";
                    ms.StartInfo.Arguments = $"-g \"C:/Program Files/CSoft/Model Studio CS/NANOPIPING/Pipe20.package\" -r Pipe";
                    ms.Start();
                    Thread.Sleep(60000);
                }
            }
            ncApp = (nanoCAD.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("nanoCADx64.Application.20.0");
            ncApp.Documents.Open(filePath);
        }

        /// <summary>
        /// Получение файла из проекта CADLib. path указывается с названием выходного файла и его типом.
        /// </summary>
        /// <param name="fileGUID">GUID файла, который требуется сохранить</param>
        /// <param name="outputPath">Путь к папке с названием файла</param>
        public void getFile(string fileGUID, string outputPath)
        {
            library.DownloadFile((int)getFileByGUID(fileGUID), outputPath);
        }

        /// <summary>
        /// Передача атрибутивной информации любому объекту базы данных.
        /// </summary>
        /// <param name="objectGUID">GUID объекта(объектом является все, кроме файлов)</param>
        /// <param name="parameterName">Название параметра</param>
        /// <param name="value">Значение параметра</param>
        public void setObjectParameter(string objectGUID, string parameterName, object value)
        {
            CLibParamDefInfo p = library.GetParamDef(parameterName);
            if (p == null)
            {
                p = new CLibParamDefInfo()
                {
                    mstrName = parameterName,
                    mstrCaption = parameterName,
                    midType = CADLibraryBase.ID_TYPE_STRING,
                    mCategories = new List<CLibParamCategoryInfo> { }
                };
                var c = library.GetParamCategoriesList();
                p.mCategories.Add(c["Системные".ToUpper()]);
                library.CreateParamDef(p);
            }
            Guid uid = Guid.Parse(objectGUID);
            library.SetObjectParameter(uid, parameterName, value.ToString(), "Parameter from 1C");
        }

        /// <summary>
        /// Передача атрибутивной информации файлу.
        /// </summary>
        /// <param name="fileGUID">GUID файла</param>
        /// <param name="parameterName">Название параметра</param>
        /// <param name="value">Значение параметра</param>
        public void setFileParameter(string fileGUID, string parameterName, object value)
        {
            CLibParamDefInfo p = library.GetParamDef(parameterName);
            if (p == null)
            {
                p = new CLibParamDefInfo()
                {
                    mstrName = parameterName,
                    mstrCaption = parameterName,
                    midType = CADLibraryBase.ID_TYPE_STRING,
                    mCategories = new List<CLibParamCategoryInfo> { }
                };
                var c = library.GetParamCategoriesList();
                p.mCategories.Add(c["Системные".ToUpper()]);
                library.CreateParamDef(p);
            }
            int? idFile = library.GetFileIdByGUID(fileGUID);
            library.DoSetFileParameter((int)idFile, parameterName, value);
        }

        /// <summary>
        /// Обновление содержимого файла по его GUID.
        /// </summary>
        /// <param name="fileGUID">GUID файла</param>
        /// <param name="filePath">Путь к файлу</param>
        public void updateFileContent(string fileGUID, string filePath)
        {
            string[] filePathMass = filePath.Split('\\');
            string newFileName = filePathMass[filePathMass.Length - 1];

            library.ReplaceFileContents(library.GetFileIdByGUID(fileGUID), filePath);
            library.UpdateFileName(library.GetFileIdByGUID(fileGUID), newFileName);
        }

        /// <summary>
        /// Переименование файла по его GUID. 
        /// </summary>
        /// <param name="newName">Новое имя файла</param>
        /// <param name="fileGUID">GUID файла</param>
        public void renameFile(string newName, string fileGUID)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"update Files set FileName = '" + newName + "' where UID = '" + fileGUID + "'";
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Удаление файла по его GUID.
        /// </summary>
        /// <param name="fileGUID">GUID файла</param>
        public void deleteFile(string fileGUID)
        {
            library.DeleteFile(library.GetFileIdByGUID(fileGUID));
        }

        /// <summary>
        /// Добавление файла (приклепление к объекту). Требуется путь.
        /// </summary>
        /// <param name="parentGUID">GUID родителя</param>
        /// <param name="filePath">Путь к файлу</param>
        /// <param name="fileCategoryCaption">Категория добавляемого файла</param>
        public string addFile(string parentGUID, string filePath, string fileCategoryCaption)
        {
            int idCategory = 1;
            Guid fileUID;

            string[] filePathMass = filePath.Split('\\');
            string fileName = filePathMass[filePathMass.Length - 1];

            if (checkCategory(fileCategoryCaption) != 0)
            {
                string sysName = getFileCategorySYSNAME(fileCategoryCaption);
                idCategory = library.GetFileCategoryId(sysName);
            }
            else
            {
                idCategory = library.AddFileCategory($"{fileCategoryCaption}_1C", fileCategoryCaption, false, false, 0);
            }

            library.DoUploadObjectFile(getObjectByGUID(parentGUID).idObject, getFileCategorySYSNAME(fileCategoryCaption), filePath, fileName);
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"select f.UID from Files f where f.idFileCategory = " + idCategory + "and f.FileName = '" + fileName + "'";
                fileUID = (Guid)cmd.ExecuteScalar();
            }
            string sfileUID = fileUID.ToString().ToUpper();
            return sfileUID;
        }

        /// <summary>
        /// Удаление объекта по его GUID.
        /// </summary>
        /// <param name="objectGUID">GUID объекта</param>
        public void deleteObject(string objectGUID)
        {
            CLibObjectInfo delObject = getObjectByGUID(objectGUID);
            library.DeleteObject(delObject);
        }

        /// <summary>
        /// Переименование объекта по его GUID.
        /// </summary>
        /// <param name="NewName">Новое имя объекта</param>
        /// <param name="objectGUID">GUID объекта</param>
        public bool renameObject(string NewName, string objectGUID)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"update Objects set Name = '" + NewName + "' where UID = '" + objectGUID + "'";
                cmd.ExecuteNonQuery();
            }
            if (getObjectByGUID(objectGUID).Name == NewName) return true;
            else return false;
        }

        /// <summary>
        /// Создание объекта "Часть документа". Если объект корневой - в таком случае не добавляем ему родителя, иначе прописываем его GUID.
        /// </summary>
        /// <param name="objectName">Имя "Части документа"</param>
        /// <param name="parentObjectGUID">GUID объекта-родителя</param>
        public string addDocPart(string objectName, string parentObjectGUID)
        {
            Dictionary<string, CLibObjectInfo> generalObjects = getGeneralObjects();
            CLibObjectInfo newObject = library.CopyObject(generalObjects["Doc_Part"], 0, false);
            library.UpdateObjectName(newObject, objectName);

            CLibObjectInfo parentObject = getObjectByGUID(parentObjectGUID);
            changeParent(newObject.idObject, parentObject.idObject);

            return library.GetObjectUIDById(newObject.idObject);
        }

        /// <summary>
        /// Создание объекта "Карточка документа". Если объект корневой - в таком случае не добавляем ему родителя, иначе прописываем его GUID.
        /// </summary>
        /// <param name="objectName">Имя "Карточки документа"</param>
        /// <param name="parentObjectGUID">GUID объекта-родителя</param>
        public string addDocCard(string objectName, string parentObjectGUID = null)
        {
            Dictionary<string, CLibObjectInfo> generalObjects = getGeneralObjects();
            CLibObjectInfo newObject = library.CopyObject(generalObjects["Doc_Card"], 0, false);
            library.UpdateObjectName(newObject, objectName);
            if (parentObjectGUID != null)
            {
                CLibObjectInfo parentObject = getObjectByGUID(parentObjectGUID);
                changeParent(newObject.idObject, parentObject.idObject);
            }
            else changeParent(newObject.idObject);

            return library.GetObjectUIDById(newObject.idObject);
        }

        /// <summary>
        /// Создание объекта "Папка". Если объект корневой - в таком случае не добавляем ему родителя, иначе прописываем его GUID.
        /// </summary>
        /// <param name="objectName">Имя "Папки"</param>
        /// <param name="parentObjectGUID">GUID объекта-родителя</param>
        public string addFolder(string objectName, string parentObjectGUID = null)
        {
            Dictionary<string, CLibObjectInfo> generalObjects = getGeneralObjects();
            CLibObjectInfo newObject = library.CopyObject(generalObjects["Folder"], 0, false);
            library.UpdateObjectName(newObject, objectName);
            if (parentObjectGUID != null)
            {
                CLibObjectInfo parentObject = getObjectByGUID(parentObjectGUID);
                changeParent(newObject.idObject, parentObject.idObject);
            }
            else changeParent(newObject.idObject);

            return library.GetObjectUIDById(newObject.idObject);
        }

        /// <summary>
        /// Создание объекта "Папки". Если объект корневой - в таком случае не добавляем ему родителя, иначе прописываем его GUID.
        /// </summary>
        /// <param name="objectName">Имя объекта "Папки"</param>
        /// <param name="parentObjectGUID">GUID объекта-родителя</param>
        public string addFolders(string objectName, string parentObjectGUID = null)
        {
            Dictionary<string, CLibObjectInfo> generalObjects = getGeneralObjects();
            CLibObjectInfo newObject = library.CopyObject(generalObjects["Folders"], 0, false);
            library.UpdateObjectName(newObject, objectName);
            if (parentObjectGUID != null)
            {
                CLibObjectInfo parentObject = getObjectByGUID(parentObjectGUID);
                changeParent(newObject.idObject, parentObject.idObject);
            }
            else changeParent(newObject.idObject);

            return library.GetObjectUIDById(newObject.idObject);
        }

        /// <summary>
        /// Подключение к БД CADLib. Ключи:
        /// "Folders",
        /// "Folder",
        /// "Doc_Card",
        /// "Doc_Part".
        /// </summary>
        /// <param name="server">Сервер, где развернута БД CADLib</param>
        /// <param name="DB">Название Базы данных</param>
        /// <param name="user">Логин пользователя, если подключаемся по аккаунту пользователя MS SQL Server</param>
        /// <param name="password">Пароль пользователя, если подключаемся по аккаунту пользователя MS SQL Server</param>
        public bool createConnection(string connectionInfo)
        {
            Dictionary<string, string> info = getCoonectInfo(connectionInfo);
            CSAppServices.DbConnectParameters dbcp = null;
            string user, password;
            if (info["db_login"] == "null" && info["db_password"] == "null")
            {
                user = null;
                password = null;
            }
            else
            {
                user = info["db_login"];
                password = info["db_password"];
            }
            if (user == null && password == null) dbcp = CSAppServices.DbConnectParameters.CreateWithOSAuth(true, LightweightDataAccess.EServerType.MSSQL, info["server"], info["db"]);
            else dbcp = CSAppServices.DbConnectParameters.CreateWithDbmsAuth(true, LightweightDataAccess.EServerType.MSSQL, (FLib.Str)info["server"], (FLib.Str)info["db"],
                                                                                (FLib.Str)user, (FLib.Str)password, null);
            bool res = library.Connect(dbcp);
            string testFolder = "";
            try
            {
                testFolder = addFolders("TestConnectFolder");
            }
            catch 
            {
                return false;
            }
            
            if (testFolder.Length != 0)
            {
                deleteObject(testFolder);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Создание модели в CADLib. Открывает форму, где требуется заполнить данные для подключения.
        /// </summary>
        public string createProjectInCADLib(string dbName)
        {
            string dbn = dbName;
            StartDialog startForm = new StartDialog(dbn);
            startForm.ShowDialog();
            if (startForm.close == false)
            {
                Process ms = new Process();
                if (startForm.connectionData["db_login"] == "null")
                {
                    ms.StartInfo.FileName = "C:/Program Files (x86)/CSoft/Model Studio CS/Viewer/bin/x64/CADLib.exe";
                    ms.StartInfo.Arguments = $"-c \"{startForm.connectionData["db_template"]}\" -s \"{startForm.connectionData["server"]}\" -d \"{startForm.connectionData["db"]}\"";
                    ms.Start();
                }
                else
                {
                    ms.StartInfo.FileName = "C:/Program Files (x86)/CSoft/Model Studio CS/Viewer/bin/x64/CADLib.exe";
                    ms.StartInfo.Arguments = $"-c \"{startForm.connectionData["db_template"]}\" -s \"{startForm.connectionData["server"]}\" -d \"{startForm.connectionData["db"]}\" -u \"{startForm.connectionData["db_login"]}\"";
                    ms.Start();
                }

                System.Text.StringBuilder connectInfo = new System.Text.StringBuilder();
                connectInfo.Append("server1C");
                connectInfo.Append(" ");
                connectInfo.Append(startForm.connectionData["server"]);
                connectInfo.Append(" ");
                connectInfo.Append("db1C");
                connectInfo.Append(" ");
                connectInfo.Append(startForm.connectionData["db"]);
                connectInfo.Append(" ");
                connectInfo.Append("template1C");
                connectInfo.Append(" ");
                connectInfo.Append(startForm.connectionData["db_template"]);
                connectInfo.Append(" ");
                connectInfo.Append("db_login");
                connectInfo.Append(" ");
                connectInfo.Append(startForm.connectionData["db_login"]);
                connectInfo.Append(" ");
                connectInfo.Append("db_password");
                connectInfo.Append(" ");
                connectInfo.Append(startForm.connectionData["db_password"]);
                return String.Join(" ", connectInfo);
            }
            else
            {
                return "";
            }
           
        }

        /// <summary>
        /// Открытие определенной модели в CADLib.
        /// </summary>
        public void openProjectInCADLib(string connectionInfo)
        {
            Dictionary<string, string> info = getCoonectInfo(connectionInfo);
            Process ms = new Process();
            ms.StartInfo.FileName = "C:/Program Files (x86)/CSoft/Model Studio CS/Viewer/bin/x64/CADLib.exe";
            ms.StartInfo.Arguments = $"-s \"{info["server"]}\" -d \"{info["db"]}\"";
            ms.Start();
        }

        /// <summary>
        /// Добавление объектов в Служебный архив 1С.
        /// </summary>
        public void objectRestoration(string objectGUID)
        {
            if ((string)library.GetObjectParameterValue(getObjectByGUID(objectGUID).idObject, "Из архива") == "False")
            {
                string parentGUID = (string)library.GetObjectParameterValue(getObjectByGUID(objectGUID).idObject, "GUID для Архива/Родительского каталога");
                changeParent(getObjectByGUID(objectGUID).idObject, getObjectByGUID(parentGUID).idObject);
                setObjectParameter(objectGUID, "Из архива", "True");
            }
        }

        /// <summary>
        /// Добавление объектов в Служебный архив 1С.
        /// </summary>
        public void addToArchive(string objectGUID)
        {
            string archiveGUID = getArchiveGUID();
            string parentGUID = library.GetObjectUIDById(getObjectByGUID(objectGUID).idParentObject);
            if (archiveGUID != parentGUID)
            {
                changeParent(getObjectByGUID(objectGUID).idObject, getObjectByGUID(archiveGUID).idObject);
                setObjectParameter(objectGUID, "GUID для Архива/Родительского каталога", parentGUID);
                setObjectParameter(objectGUID, "Из архива", "False");
            }
        }

        #endregion

        #region Функции для добавления нового каталога и остальных действий (старые)

        /// <summary>
        /// Создание каталога. Возвращает объект каталога. (решить насчет фильтра)
        /// </summary>
        public CLibCatalogFilterItem createCatalog(string catalogName)
        {
            List<CLibFilterItem> filterNew = new List<CLibFilterItem> { };
            CLibCatalogFilterItem catalog = library.CreateCatalog(null, catalogName, filterNew, 0);
            return catalog;
        }

        /// <summary>
        /// Создание проекта. Возвращает ID проекта.
        /// </summary>
        public int? createProject(string projectName, string projectCategoryName)
        {
            int prjCategory = getObjectCategory("project", projectCategoryName);
            library.CreateObjectRecord(projectName, prjCategory, 0, 0, 0, 0);
            int? idProject = getObject(projectName, prjCategory, null);
            return idProject;
        }

        /// <summary>
        /// Создание папки в конкретном каталоге/проекте.
        /// </summary>
        public int? createFolder(string folderName, int folderCategory, int idParent)
        {
            int temp = library.CreateObjectRecord(folderName, folderCategory, (int)idParent, 0, 0, 0);
            return temp;
        }

        /// <summary>
        /// Создание файла в папке. В CADLib отображает во вкладке вложения.
        /// Присутствует возможность сразу задать файлу GUID. 
        /// </summary>
        public void createFileInFolder(string folderName, int catCategory, int idFolderParent, string filePath, string fileName, string fileCategory, Guid guid)
        {
            int? idFolder = getObject(folderName, catCategory, idFolderParent);
            library.DoUploadObjectFile((int)idFolder, fileCategory, filePath, fileName, guid);
        }

        /// <summary>
        /// Переименование папки.
        /// </summary>
        public void renameFolder(string oldFolderName, string newFolderName, int folderCategory, int idParent)
        {
            int? idFolder = getObject(oldFolderName, folderCategory, idParent);
            CLibObjectInfo folder = library.GetLibraryObject((int)idFolder);
            library.UpdateObjectName(folder, newFolderName);
        }

        /// <summary>
        /// Удаление папки.
        /// </summary>
        public void deleteFolder(string folderName, int catCategory, int idParent)
        {
            int? idFolder = getObject(folderName, catCategory, idParent);
            CLibObjectInfo folderDel = library.GetLibraryObject((int)idFolder);
            library.DeleteObject(folderDel);
        }

        /// <summary>
        /// Переименование файла по GUID.
        /// </summary>
        public void renameFileByGUID(Guid guid, string newFileName)
        {
            int? idFile = getFileByGUID(guid.ToString());
            library.UpdateFileName((int)idFile, newFileName);
        }

        /// <summary>
        /// Удаление файла по GUID.
        /// </summary>
        public void deleteFileByGUID(Guid guid)
        {
            int? idFile = getFileByGUID(guid.ToString());
            CLibObjectInfo fileDel = library.GetLibraryObject((int)idFile);
            library.DeleteObject(fileDel);
        }

        #endregion ()

        #region Вспомогательные функции

        /// <summary>
        /// Проверка существования категории с данным именем(caption)
        /// </summary>
        private Dictionary<string, string> getCoonectInfo(string connectionInfo)
        {
            Dictionary<string, string> res = new Dictionary<string, string>(3);

            string[] info = connectionInfo.Split();
            StringBuilder infoMass = new StringBuilder();
            infoMass.Append(info[1]);
            res.Add("server", String.Join("", infoMass));
            infoMass.Clear();

            for (int i = 0; i < info.Length; i++)
            {
                if (info[i] == "db1C")
                {
                    i++;
                    while (info[i] != "template1C")
                    {
                        infoMass.Append(info[i]);
                        if (info[i + 1] != "template1C") infoMass.Append(" ");
                        i++;
                    }
                    break;
                }
            }
            res.Add("db", String.Join("", infoMass));
            infoMass.Clear();

            for (int i = 0; i < info.Length; i++)
            {
                if (info[i] == "template1C")
                {
                    i++;
                    while (info[i] != "db_login")
                    {
                        infoMass.Append(info[i]);
                        if (info[i + 1] != "db_login") infoMass.Append(" ");
                        i++;
                    }
                    break;
                }
            }
            res.Add("template", String.Join("", infoMass));
            infoMass.Clear();

            for (int i = 0; i < info.Length; i++)
            {
                if (info[i] == "db_login")
                {
                    i++;
                    while (info[i] != "db_password")
                    {
                        infoMass.Append(info[i]);
                        if (info[i + 1] != "db_password") infoMass.Append(" ");
                        i++;
                    }
                    break;
                }
            }
            res.Add("db_login", String.Join("", infoMass));
            infoMass.Clear();

            for (int i = 0; i < info.Length; i++)
            {
                if (info[i] == "db_password")
                {
                    i++;
                    while (i != info.Length)
                    {
                        infoMass.Append(info[i]);
                        if (i != info.Length-1) infoMass.Append(" ");
                        i++;
                    }
                    break;
                }
            }
            res.Add("db_password", String.Join("", infoMass));
            infoMass.Clear();

            //MessageBox.Show(res["server"]);
            //MessageBox.Show(res["db"]);
            //MessageBox.Show(res["server"]);

            return res;
        }

        /// <summary>
        /// Запуск приложения с указанием открываемого файла. 
        /// </summary>
        private bool OpenFile(string fileGUID, string fileName, string programPath)
        {
            int idFile = library.GetFileIdByGUID(fileGUID);
            string filePath = Path.Combine(Environment.GetEnvironmentVariable("TEMP"), fileName);

            library.DownloadFile((int)idFile, "C:/Users/shevchenko/Desktop/TestNanoCad/testDwgFile");

            Process.Start(programPath, '"' + filePath + '"');
            return true;
        }

        /// <summary>
        /// Получение GUID Служебного архива в базе данных.
        /// </summary>
        private string getArchiveGUID()
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"select o.UID from Objects o where o.idParentObject is NULL and o.Name = 'Служебный архив 1С'";
                return (string)cmd.ExecuteScalar().ToString();
            }
        }

        /// <summary>
        /// Проверка существования категории с данным именем(caption)
        /// </summary>
        private int checkCategory(string categoryName)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"SELECT COUNT(1) FROM FileCategories f WHERE f.Caption = '" + categoryName + "'";
                return (int)cmd.ExecuteScalar();
            }
        }

        /// <summary>
        /// Получение SYS_NAME категории.
        /// </summary>
        private string getFileCategorySYSNAME(string categoryCaption)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"select fc.SysName from FileCategories fc where fc.Caption = '" + categoryCaption + "'";
                return (string)cmd.ExecuteScalar();
            }
        }

        /// <summary>
        /// Получение GUID файла.
        /// </summary>
        private string getFileGUID(int idCategory, string fileName)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"select f.UID from Files where FileName = '" + fileName + "' and idFileCategory = '" + idCategory + "'";
                return (string)cmd.ExecuteScalar();
            }
        }

        /// <summary>
        /// Добавление объекту GUID из 1C.
        /// </summary>
        private void setObjectGUID(int idObject, string newObjectGUID)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"update Objects set UID = '" + newObjectGUID + "' where idObject = '" + idObject + "'";

                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Получение объекта по GUID. На выходе объект CLibObjectInfo. 
        /// </summary>
        private CLibObjectInfo getObjectByGUID(string objectGUID)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"select o.idObject from Objects o where UID = '" + objectGUID + "'";
                int? idObject = (int?)cmd.ExecuteScalar();

                return library.GetLibraryObject((int)idObject);
            }
        }

        /// <summary>
        /// Получение нужных объектов: "Папки", "Папка", "Карточка документа", "Часть документа"
        /// </summary>
        private Dictionary<string, CLibObjectInfo> getGeneralObjects()
        {
            int? id_foldersEmpty = getObjectByName("NeedFor1C_2");
            CLibObjectInfo objectFoldersEmpty = library.GetLibraryObject((int)id_foldersEmpty);
            int? id_folders = getObjectByName("NeedFor1C");
            CLibObjectInfo objectFolders = library.GetLibraryObject((int)id_folders);
            int? id_folder = getObjectByName("Папка1C");
            CLibObjectInfo objectFolder = library.GetLibraryObject((int)id_folder);
            int? id_docCard = getObjectByName("Карточка документа1C");
            CLibObjectInfo docCard = library.GetLibraryObject((int)id_docCard);
            int? id_docPart = getObjectByName("Часть документа1C");
            CLibObjectInfo docPart = library.GetLibraryObject((int)id_docPart);

            Dictionary<string, CLibObjectInfo> generalObjects = new Dictionary<string, CLibObjectInfo>(4);
            generalObjects.Add("Folders", objectFoldersEmpty);
            generalObjects.Add("Folder", objectFolder);
            generalObjects.Add("Doc_Card", docCard);
            generalObjects.Add("Doc_Part", docPart);

            return generalObjects;
        }

        /// <summary>
        /// Получение ID объекта по его имени
        /// </summary>
        private int? getObjectByName(string objectName)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"select idObject from Objects where Name = '" + objectName + "'";
                int? res = (int?)cmd.ExecuteScalar();
                return res;
            }
        }

        /// <summary>
        /// Смена родительского объекта. Если newIdObjectParent оставить равным 0, тогда объект сместиться в родительскую папку "Документы проекта".
        /// </summary>
        private void changeParent(int idCurrentObject, int newIdObjectParent = 0)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                if (newIdObjectParent == 0) cmd.CommandText = @"update Objects set idParentObject = NULL where idObject = " + idCurrentObject;
                else cmd.CommandText = @"update Objects set idParentObject = '" + newIdObjectParent + "' where idObject = " + idCurrentObject;

                cmd.ExecuteNonQuery();
            }
        }


        /// <summary>
        /// Создание категории. Новая категория высветиться в CADLib при обновлении БД в приложении.
        /// </summary>
        private void createCategory(string categoryName, string categoryCaption)
        {
            int nCat = library.CreateCategory(categoryName, categoryCaption);
            library.UpdateCategoriesList();
        }

        /// <summary>
        /// Возвращает ID папки с заданным именем в проекте с заданным именем. Создаёт папку и проект, если они не существуют.
        /// </summary>
        /// <param name="prjName">Имя проекта</param>
        /// <param name="folderName">Имя папки</param>
        /// <returns></returns>
        private int? getFolder(string prjName, string folderName)
        {
            int prjCategory = getObjectCategory("project", "Проект");
            int? idProject = getObject(prjName, prjCategory, null);

            int folderCategory = getObjectCategory("projectFolder", "Папка проекта");
            return getObject(folderName, folderCategory, idProject);
        }

        /// <summary>
        /// Возвращает ID объекта по его имени, категории и родительскому объекту. Имя должно быть уникальным в пределах категории и родительского объекта.
        /// </summary>
        /// <param name="prjName">Имя проекта</param>
        /// <param name="objCategory">ID категории проекта</param>
        /// <returns>ID проекта</returns>
        private int? getObject(string objName, int objCategory, int? idParent)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"select idObject from Objects where lower(Name) = lower('" + objName + "') and idObjectCategory = " + objCategory;
                if (idParent.HasValue)
                {
                    cmd.CommandText = cmd.CommandText + " and idParentObject = " + idParent;
                }
                return (int?)cmd.ExecuteScalar();
            }

        }

        /// <summary>
        /// Возвращает ID категории с заданным системным именем. При отсутствии тековой создаёт категорию с заданным системным именем и пользовательским именем.
        /// </summary>
        /// <param name="name">Системное имя</param>
        /// <param name="caption">Пользовательское имя</param>
        /// <returns>ID категории</returns>
        private int getObjectCategory(string name, string caption)
        {
            int id = library.GetObjectCategoryId(name);
            if (id <= 0)
            {
                id = library.CreateCategory(name, caption);
            }
            return id;
        }

        /// <summary>
        /// Возвращает ID файла по его GUID.
        /// </summary>
        /// <param name="idObject">ID объекта</param>
        /// <param name="fileName">Имя файла</param>
        /// <returns></returns>
        private int? getFileByGUID(string fileGUID)
        {
            using (var cmd = library.Connection.CreateCommand())
            {
                cmd.CommandText = @"select f.idFile from Files f where UID = '" + fileGUID + "'";

                return (int?)cmd.ExecuteScalar();
            }
        }
        #endregion
    }
}
