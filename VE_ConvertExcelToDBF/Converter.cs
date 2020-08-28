using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SocialExplorer.IO.FastDBF;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace VE_ConvertExcelToDBF
{
    public class Abonent
    {
        public string Account;
        public string Fio;
        //public string Address;
        //public string Alias;
    }

    public class RegistryInfo
    {
        public DateTime Now;
        public List<Abonent> Abonents = new List<Abonent>();
        public string FileName;
    }

    public class Converter
    {
        public List<Abonent> GetFizData(string path, string fileExtension)
        {
            List<Abonent> abonents = new List<Abonent>();

            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);

            IWorkbook workBook;
            if (fileExtension.ToLower().Equals(".xls"))
            {
                workBook = new HSSFWorkbook(fs);
            }
            else
            {
                workBook = new XSSFWorkbook(fs);
            }

            ISheet sheet = workBook.GetSheetAt(0);
            for (int row_num = 3; row_num <= sheet.LastRowNum; row_num++)
            {
                var row = sheet.GetRow(row_num);

                if (row != null)
                {
                    ulong Account = 0;
                    var cell = row?.GetCell(3)?.ToString();
                    if (string.IsNullOrEmpty(cell))
                    {
                        Console.WriteLine("Не попадёт в dbf row.RowNum: " + row.RowNum);
                        continue;
                    }

                    if (!ulong.TryParse(cell, out Account)) //Ожидает, что второй столбец - л/с
                    {
                        Console.WriteLine("Не попадёт в dbf: " + row.GetCell(3).ToString());
                        continue;
                    }

                    abonents.Add(new Abonent
                    {
                        Fio = row.GetCell(2).ToString(),
                        Account = row.GetCell(3).ToString()
                    });
                }
            }

            return abonents;
        }
        //public List<Abonent> GetUrData(string path, string alias, string fileExtension)
        //{
        //    List<Abonent> abonents = new List<Abonent>();

        //    FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);

        //    IWorkbook workBook;
        //    if (fileExtension.ToLower().Equals(".xls"))
        //    {
        //        workBook = new HSSFWorkbook(fs);
        //    }
        //    else
        //    {
        //        workBook = new XSSFWorkbook(fs);
        //    }

        //    ISheet sheet = workBook.GetSheetAt(0);

        //    for (int row_num = 0; row_num <= sheet.LastRowNum; row_num++)
        //    {
        //        var row = sheet.GetRow(row_num);

        //        if (row != null)
        //        {
        //            ulong Account = 0;
        //            var cell = row?.GetCell(1)?.ToString();
        //            if (string.IsNullOrEmpty(cell))
        //            {
        //                Console.WriteLine("Не попадёт в dbf row.RowNum: " + row.RowNum);
        //                continue;
        //            }

        //            if (!ulong.TryParse(cell, out Account)) //Ожидает, что второй столбец - л/с
        //            {
        //                Console.WriteLine("Не попадёт в dbf: " + row.GetCell(1).ToString());
        //                continue;
        //            }

        //            abonents.Add(new Abonent
        //            {
        //                Account = row.GetCell(1).ToString(),
        //                Fio = row.GetCell(2).ToString(),
        //                Address = row.GetCell(3).ToString(),
        //                Alias = alias
        //            });
        //        }
        //    }

        //    return abonents;
        //}

        public List<Abonent> GetFizDataFromDbf(string path)
        {
            List<Abonent> abonents = new List<Abonent>();
            var dbfFileToRead = new DbfFile();
            dbfFileToRead.Open(path, FileMode.Open);

            DbfRecord record = new DbfRecord(dbfFileToRead.Header);
            while (dbfFileToRead.ReadNext(record))
            {
                abonents.Add(
                    new Abonent()
                    {
                        Fio = record[1],
                        Account = record[2]
                    }
                );
            }
            dbfFileToRead.Close();

            return abonents;
        }

        public string CreateDbfFile(RegistryInfo registryInfo)
        {
            try
            {
                var workDirectory = Path.Combine(Directory.GetCurrentDirectory(), @"Abonents\" + registryInfo.Now.ToString("dd.MM.yyyy"));
                var dbfFilePath = Path.Combine(workDirectory, $"{registryInfo.FileName}" + ".dbf");

                if (!Directory.Exists(workDirectory))
                    Directory.CreateDirectory(workDirectory);

                if (File.Exists(dbfFilePath))
                    File.Delete(dbfFilePath);

                var dbfFile = new DbfFile();
                dbfFile.Create(dbfFilePath);
                dbfFile.Header.Unlock();
                dbfFile.Header.AddColumn("FIO", DbfColumn.DbfColumnType.Character, 100, 0);
                dbfFile.Header.AddColumn("ACCOUNT", DbfColumn.DbfColumnType.Character, 20, 0);

                var record = new DbfRecord(dbfFile.Header, Encoding.GetEncoding(1251)) { };
                foreach (Abonent row in registryInfo.Abonents)
                {
                    record[0] = row.Fio;
                    record[1] = row.Account;
                    dbfFile.Write(record, true);
                }
                dbfFile.Close();
                return dbfFilePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error in CreateDbfFile: {ex}");
            }
        }
    }
}
