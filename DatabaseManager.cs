
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;



namespace lab5
{
    internal class DatabaseManager
    {
        private List<Paintings> _listPaintings;
        private List<Artists> _listArtists;
        private List<Styles> _listStyles;

        public List<Paintings> ListPaintings
        {
            get { return _listPaintings; }
            set { _listPaintings = value; }
        }

        public List<Artists> ListArtists
        {
            get { return _listArtists; }
            set { _listArtists = value; }
        }

        public List<Styles> ListStyle
        {
            get { return _listStyles; }
            set { _listStyles = value; }
        }

        public DatabaseManager(List<Paintings> listPaintings, List<Artists> listArtists, List<Styles> listStyles)
        {
            ListPaintings = listPaintings;
            ListArtists = listArtists;
            ListStyle = listStyles;
        }

        // 2 Пункт
        public void ReadFromExcel()
        {
            Validator validator = new Validator();
            int sheetNumber;
            Console.WriteLine("Какую базу данных хотите прочитать из Excel:");
            Console.WriteLine("1 - Картины");
            Console.WriteLine("2 - Художники");
            Console.WriteLine("3 - Стили живописи");
            sheetNumber = Convert.ToInt32(validator.check_int(Console.ReadLine()));
            switch (sheetNumber)
            {
                case 1:
                    {
                        Console.WriteLine("Картины:");
                        foreach (Paintings painting in ListPaintings)
                        {
                            Console.WriteLine(painting);
                        }
                        break;
                    }
                case 2:
                    {
                        Console.WriteLine("Художники:");
                        foreach (Artists artist in ListArtists)
                        {
                            Console.WriteLine(artist);
                        }
                        break;
                    }
                case 3:
                    {
                        Console.WriteLine("Стили жиповиси:");
                        foreach (Styles style in ListStyle)
                        {
                            Console.WriteLine(style);
                        }
                        break;
                    }
            }
        }

        // 3 Пункт
        public void DeleteElements()
        {
            Validator validator = new Validator();
            int sheetNumber;
            int itemNum;
            Console.WriteLine("Напишите базу данных из которой вы хотите удалить элемент: ");
            Console.WriteLine("1 - Картины");
            Console.WriteLine("2 - Художники");
            Console.WriteLine("3 - Стили живописи");
            sheetNumber = Convert.ToInt32(validator.check_int(Console.ReadLine()));
            Console.WriteLine();
            Console.Write("Напишите номер элемента, который вы хотите удалить: ");
            itemNum = Convert.ToInt32(validator.check_int(Console.ReadLine()));
            Console.WriteLine();

            bool found = false;


            switch (sheetNumber)
            {
                case 1:
                    {
                        
                        for (int i = 0; i < ListPaintings.Count; ++i)
                        {
                            if (ListPaintings[i].ID == itemNum)
                            {
                                found = true;
                                this.ListPaintings.RemoveAt(i);
                                Console.WriteLine("Из базы данных Картины удален элемент с ID = {0}", itemNum);
                                break;
                            }
                        }

                        if (found == false)
                        {
                            Console.WriteLine("Не найдено ID = {0} в базе данных Картины", itemNum);
                            break;
                        }
                        
                        break;
                    }
                case 2:
                    {
                        for (int i = 0; i < ListArtists.Count; ++i)
                        {
                            if (ListArtists[i].ID == itemNum)
                            {
                                found = true;
                                this.ListArtists.RemoveAt(i);
                                Console.WriteLine("Из базы данных Художники удален элемент с ID = {0}", itemNum);
                                break;
                            }
                        }

                        if (found == false)
                        {
                            Console.WriteLine("Не найдено ID = {0} в базе данных Художники", itemNum);
                            break;
                        }
                        break;
                    }
                case 3:
                    {
                        for (int i = 0; i < ListStyle.Count; ++i)
                        {
                            if (ListStyle[i].ID == itemNum)
                            {
                                found = true;
                                this.ListStyle.RemoveAt(i);
                                Console.WriteLine("Из базы данных Стиль удален элемент с ID = {0}", itemNum);
                                break;
                            }
                        }

                        if (found == false)
                        {
                            Console.WriteLine("Не найдено ID = {0} в базе данных Стиль", itemNum);
                            break;
                        }
                        break;
                    }
            }
        }

        // 4 пункт
        public void AddElement()
        {
            Validator validator = new Validator();
            int sheetNumber;
            Console.WriteLine("Напишите базу данных в которую вы хотите добавить элемент: ");
            Console.WriteLine("1 - Картины");
            Console.WriteLine("2 - Художники");
            Console.WriteLine("3 - Стили живописи");
            sheetNumber = Convert.ToInt32(validator.check_int(Console.ReadLine()));
            Console.WriteLine();
            switch (sheetNumber)
            {
                case 1:
                    {
                        Paintings painting = new Paintings();
                        Console.Write("Напишите название Картины: ");
                        string namePainting = Console.ReadLine();
                        Console.Write("Напишите ID Художника, к которому относится эта картина: ");

                        // Если введен неизвестный ID художника
                        int idAtrist = Convert.ToInt32(validator.check_int(Console.ReadLine()));
                        bool foundA = false;
                        for (int i = 0; i < ListArtists.Count; ++i)
                        {
                            if (ListArtists[i].ID == idAtrist)
                            {
                                foundA = true;
                                break;
                            }
                        }
                        if (foundA == false)
                        {
                            Console.WriteLine("Не найдено ID = {0} в базе данных Художники", idAtrist);
                            break;
                        }

                        Console.Write("Напишите часть эрмитажа, в которой находится эта картина: ");
                        int partOfHermitage = Convert.ToInt32(validator.check_int(Console.ReadLine()));
                        Console.Write("Напишите дату создания этой картины: ");
                        int year = Convert.ToInt32(validator.check_int(Console.ReadLine()));

                        Console.Write("Напишите ID стиля, к которому относится эта картина: ");
                        int idStyle = Convert.ToInt32(validator.check_int(Console.ReadLine()));
                        bool foundS = false;
                        for (int i = 0; i < ListStyle.Count; ++i)
                        {
                            if (ListStyle[i].ID == idStyle)
                            {
                                foundS = true;
                                break;
                            }
                        }
                        if (foundS == false)
                        {
                            Console.WriteLine("Не найдено ID = {0} в базе данных Стиль", idStyle);
                            break;
                        }

                        int newID = 1;
                        foreach (Paintings paint in ListPaintings)
                        {
                            if (paint.ID >= newID)
                                newID = paint.ID + 1;
                        }


                        painting.ID = newID;
                        painting.NamePainting = namePainting;
                        painting.IdArtist = idAtrist;
                        painting.PartOdHermitage = partOfHermitage;
                        painting.Year = year;
                        painting.IdStyle = idStyle;

                        ListPaintings.Add(painting);
                        break;
                    }
                case 2:
                    {
                        Artists artist = new Artists();
                        Console.Write("Напишите имя и фамилию художника: ");
                        string name = Console.ReadLine();

                        int newID = 1;
                        foreach (Artists art in ListArtists)
                        {
                            if (art.ID >= newID)
                                newID = art.ID + 1;
                        }

                        artist.ID = newID;
                        artist.Name = name;
                        ListArtists.Add(artist);
                        break;
                    }
                case 3:
                    {
                        Styles style = new Styles();
                        Console.Write("Напишите название стиля жанра: ");
                        string paintingStyle = Console.ReadLine();

                        int newID = 1;
                        foreach (Styles st in ListStyle)
                        {
                            if (st.ID >= newID)
                                newID = st.ID + 1;
                        }

                        style.ID = newID;
                        style.PaintingStyle = paintingStyle;
                        ListStyle.Add(style);
                        break;
                    }
            }
        }


        // 5 Пункт
        // Пример из файла
        public void QueryExample()
        {
            int query = (from painting in ListPaintings
                         where painting.PartOdHermitage == 2
                         group painting by painting.IdArtist into artistGroup
                         where artistGroup.Count() > 5
                         select artistGroup.Key).Count();

            Console.WriteLine("Количество художников с более чем 5 картинами во 2-й части Эрмитажа: {0}", query);
        }

        // Найти средний год создания картин указанного художника
        public void Query2()
        {
            //Адам, Альбрехт
            Console.WriteLine("Введите имя художника: ");
            string nameArtist = Console.ReadLine();

            var query = (from painting in ListPaintings
                         join artist in ListArtists on painting.IdArtist equals artist.ID
                         where artist.Name == nameArtist
                         select painting.Year).Average();
            Console.WriteLine("Средний год создания картин художника {0}: {1:F2}", nameArtist, query);
        }

        // Получить перечень картин с названиями художников и стилей
        public void Query3()
        {
            var query = (from painting in ListPaintings
                         join artist in ListArtists on painting.IdArtist equals artist.ID
                         join style in ListStyle on painting.IdStyle equals style.ID
                         select new { painting.NamePainting, artist.Name, style.PaintingStyle });
            Console.WriteLine("перечень картин с названиями художников и стилей:");
            foreach (var item in query)
            {
                Console.WriteLine("Название: {0}\nАвтор: {1}\nСтиль: {2}", item.NamePainting, item.Name, item.PaintingStyle);
                Console.WriteLine();
            }
        }

        // Количество картин каждого художника по стилям
        public void Query4()
        {
            var query = (from painting in ListPaintings
                         join artist in ListArtists on painting.IdArtist equals artist.ID
                         join style in ListStyle on painting.IdStyle equals style.ID
                         group painting by new { artist.Name, style.PaintingStyle } into g
                         where g.Count() > 1
                         orderby g.Count() descending
                         select new
                         {
                             Artist = g.Key.Name,
                             Style = g.Key.PaintingStyle,
                             Count = g.Count(),
                         });
            Console.WriteLine("Художники и их стили:");
            foreach (var item in query)
            {
                Console.WriteLine("{0} - {1}: {2} картин", item.Artist, item.Style, item.Count);
                Console.WriteLine();
            }
        }

        // 6 пункт 
        public void SaveToExcel()
        {
            try
            {
                IWorkbook workbook = new HSSFWorkbook();

               
                ISheet paintingSheet = workbook.CreateSheet("Картины");
                UpdatePaintingSheet(paintingSheet);

               
                ISheet artistsSheet = workbook.CreateSheet("Художники");
                UpdateArtistsSheet(artistsSheet);

                
                ISheet stylesSheet = workbook.CreateSheet("Стиль");
                UpdateStylesSheet(stylesSheet);

          
                using (var fileStream = new FileStream("LR5-var11.xls", FileMode.Create))
                {
                    workbook.Write(fileStream);
                }

                Console.WriteLine("Изменения успешно сохранены в файл!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при сохранении: {0}", ex.Message);
            }
        }

        private void UpdatePaintingSheet(ISheet sheet)
        {
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("ID");
            headerRow.CreateCell(1).SetCellValue("Название");
            headerRow.CreateCell(2).SetCellValue("ID Художника");
            headerRow.CreateCell(3).SetCellValue("Часть эрмитажа");
            headerRow.CreateCell(4).SetCellValue("Год");
            headerRow.CreateCell(5).SetCellValue("ID стиля");

            for (int i = 0; i < ListPaintings.Count; i++)
            {
                Paintings painting = ListPaintings[i];
                IRow row = sheet.CreateRow(i + 1);

                row.CreateCell(0).SetCellValue(painting.ID);
                row.CreateCell(1).SetCellValue(painting.NamePainting);
                row.CreateCell(2).SetCellValue(painting.IdArtist);
                row.CreateCell(3).SetCellValue(painting.PartOdHermitage);
                row.CreateCell(4).SetCellValue(painting.Year);
                row.CreateCell(5).SetCellValue(painting.IdStyle);
            }
        }

        private void UpdateArtistsSheet(ISheet sheet)
        {
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("ID");
            headerRow.CreateCell(1).SetCellValue("Имя");

            for (int i = 0; i < ListArtists.Count; i++)
            {
                Artists artist = ListArtists[i];
                IRow row = sheet.CreateRow(i + 1);

                row.CreateCell(0).SetCellValue(artist.ID);
                row.CreateCell(1).SetCellValue(artist.Name);
            }
        }

        private void UpdateStylesSheet(ISheet sheet)
        {
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("ID");
            headerRow.CreateCell(1).SetCellValue("Название");

            for (int i = 0; i < ListStyle.Count; i++)
            {
                Styles style = ListStyle[i];
                IRow row = sheet.CreateRow(i + 1);

                row.CreateCell(0).SetCellValue(style.ID);
                row.CreateCell(1).SetCellValue(style.PaintingStyle);
            }
        }
    }
}
