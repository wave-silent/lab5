using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

namespace lab5
{
    internal class Program
    {

        static void Main(string[] args)
        {
            Validator validator = new Validator();
            int num1;
            Console.Write("Введите номер задания из списка \"1 2 3 4 5 6\": ");
            num1 = Convert.ToInt32(validator.check_int(Console.ReadLine()));

            switch (num1)
            {
                case 1:
                    {
                        Paintings painting = new Paintings();
                        Artists artists = new Artists();
                        Styles styles = new Styles();

                        painting.UploadPaintingsFromExcel();
                        artists.UploadArtistsFromExcel();
                        styles.UploadStylesFromExcel();

                        break;
                    }
                case 2:
                    {
                        Paintings painting = new Paintings();
                        Artists artists = new Artists();
                        Styles styles = new Styles();

                        DatabaseManager database = new DatabaseManager(
                            painting.UploadPaintingsFromExcel(),
                            artists.UploadArtistsFromExcel(),
                            styles.UploadStylesFromExcel()
                        );

                        database.ReadFromExcel();
                        break;
                    }
                case 3:
                    {
                        Paintings painting = new Paintings();
                        Artists artists = new Artists();
                        Styles styles = new Styles();

                        DatabaseManager database = new DatabaseManager(
                            painting.UploadPaintingsFromExcel(),
                            artists.UploadArtistsFromExcel(),
                            styles.UploadStylesFromExcel()
                        );

                        // макс кол-во обьектов в списке 698, так как отсутсвуте 1 строка и 35 ID
                        Console.WriteLine(database.ListPaintings[697]);
                        
                        database.DeleteElements();

                        Console.WriteLine(database.ListPaintings[696]);
                        break;
                    }
                case 4:
                    {
                        Paintings painting = new Paintings();
                        Artists artists = new Artists();
                        Styles styles = new Styles();

                        DatabaseManager database = new DatabaseManager(
                            painting.UploadPaintingsFromExcel(),
                            artists.UploadArtistsFromExcel(),
                            styles.UploadStylesFromExcel()
                        );

                        database.AddElement();
                       
                        Console.WriteLine(database.ListArtists[290]);
                        break;
                    }
                case 5:
                    {
                        Paintings painting = new Paintings();
                        Artists artists = new Artists();
                        Styles styles = new Styles();

                        DatabaseManager database = new DatabaseManager(
                            painting.UploadPaintingsFromExcel(),
                            artists.UploadArtistsFromExcel(),
                            styles.UploadStylesFromExcel()
                        );

                        Console.WriteLine("Какой запрос хотите выполнить: ");
                        int q = Convert.ToInt32(validator.check_int(Console.ReadLine()));
                        switch (q)
                        {
                            case 1:
                                {
                                    database.QueryExample();
                                    break;
                                }
                            case 2:
                                {
                                    database.Query2();
                                    break;
                                }
                            case 3:
                                {
                                    database.Query3();
                                    break;
                                }
                            case 4:
                                {
                                    database.Query4();
                                    break;
                                }
                        }                 
                        break;
                    }
                case 6:
                    {
                        Paintings painting = new Paintings();
                        Artists artists = new Artists();
                        Styles styles = new Styles();

                        DatabaseManager database = new DatabaseManager(
                            painting.UploadPaintingsFromExcel(),
                            artists.UploadArtistsFromExcel(),
                            styles.UploadStylesFromExcel()
                        );


                        database.AddElement();

                        database.SaveToExcel();
                        break;
                    }
            }
        }
    }
}
