using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VE_ConvertExcelToDBF
{
    public class Program
    {
        static void Main(string[] args)
        {
            var converter = new Converter();
            RegistryInfo registryInfo = new RegistryInfo();
            registryInfo.Now = DateTime.Now;

            //YssykKulAbonentsFiz.dbf
            //YssykKulAbonentsUr.dbf
            //NarynAbonentsFiz.dbf
            //NarynAbonentsUr.dbf

            /*
            registryInfo.FileName = "YssykKulAbonentsFiz";
            var files = Directory.GetFiles($@"{Directory.GetCurrentDirectory()}/Ыссык - Куль FizLitsa");
            for (int i = 0; i<7; i++)
            {
                var alias = getAliasOfRes(Path.GetFileNameWithoutExtension(files[i]));
                var Abonents = converter.GetFizData(files[i], alias, Path.GetExtension(files[i]));

                registryInfo.Abonents.AddRange(Abonents);
            }
            */

            /*registryInfo.FileName = "YssykKulAbonentsUr";
            var files = Directory.GetFiles($@"{Directory.GetCurrentDirectory()}/Ыссык - Куль UrLitsa");
            for (int i = 0; i < 7; i++)
            {
                var alias = getAliasOfRes(Path.GetFileNameWithoutExtension(files[i]));
                var Abonents = converter.GetUrData(files[i], alias, Path.GetExtension(files[i]));

                registryInfo.Abonents.AddRange(Abonents);
            }*/

            registryInfo.FileName = "BishkekFilial_25.08.2020";
            var files = Directory.GetFiles($@"{Directory.GetCurrentDirectory()}/TrustUnion");

                List<Abonent> Abonents;
                if (Path.GetExtension(files[0]) == ".dbf")
                {
                    Abonents = converter.GetFizDataFromDbf(files[0]);
                }
                else
                {
                    Abonents = converter.GetFizData(files[0], Path.GetExtension(files[0]));
                }
                registryInfo.Abonents.AddRange(Abonents);

            //registryInfo.FileName = "TokmokFilial_25.08.2020";
            //var files = Directory.GetFiles($@"{Directory.GetCurrentDirectory()}/TrustUnion");

            //List<Abonent> Abonents;
            //if (Path.GetExtension(files[0]) == ".dbf")
            //{
            //    Abonents = converter.GetFizDataFromDbf(files[0]);
            //}
            //else
            //{
            //    Abonents = converter.GetFizData(files[0], Path.GetExtension(files[0]));
            //}
            //registryInfo.Abonents.AddRange(Abonents);



            /*registryInfo.FileName = "NarynAbonentsUr";
            var files = Directory.GetFiles($@"{Directory.GetCurrentDirectory()}/Нарын UrLitsa");
            for (int i = 0; i < 7; i++)
            {
                var alias = getAliasOfRes(Path.GetFileNameWithoutExtension(files[i]));
                var Abonents = converter.GetUrData(files[i], alias, Path.GetExtension(files[i]));

                registryInfo.Abonents.AddRange(Abonents);
            }*/




            converter.CreateDbfFile(registryInfo);
            Console.ReadKey();
        }

        //static string getAliasOfRes(string alias)
        //{
        //    switch (alias)
        //    {
        //        case "KarakolFiz":
        //        case "AkSuuFiz":
        //        case "TupFiz":
        //        case "TonFiz":
        //        case "JetiOguzFiz":
        //        case "CholponAtaFiz":
        //        case "BalykchyFiz":

        //        case "AkTalaaFiz":
        //        case "AtBashyFiz":
        //        case "JumgalFiz":
        //        case "KochkorFiz":
        //        case "NarynFiz":
        //        case "ToguzToroFiz":
        //        case "TyanShyanFiz":

        //        case "KarakolUr":
        //        case "AkSuuUr":
        //        case "TupUr":
        //        case "TonUr":
        //        case "JetiOguzUr":
        //        case "CholponAtaUr":
        //        case "BalykchyUr":

        //        case "AkTalaaUr":
        //        case "AtBashyUr":
        //        case "JumgalUr":
        //        case "KochkorUr":
        //        case "NarynUr":
        //        case "ToguzToroUr":
        //        case "TyanShyanUr":
        //            return alias;
        //        default:
        //            throw new Exception("Alias not found!");
        //    }
        //}
    }
}
