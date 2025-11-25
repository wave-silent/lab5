using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5
{
    internal class Paintings
    {
        private int _id;
        private string _namePainting;
        private int _idAtrist;
        private int _partOfHermitage;
        private int _year;
        private int _idStyle;
        private List<Paintings> _listPaintings;

        public int ID
        {
            get { return _id; }
            set { _id = value; }
        }
        public string NamePainting
        {
            get { return _namePainting; }
            set { _namePainting = value; }
        }
        public int IdArtist
        {
            get { return _idAtrist; }
            set { _idAtrist = value; }
        }
        public int PartOdHermitage
        {
            get { return _partOfHermitage; }
            set { _partOfHermitage = value; }
        }
        public int Year
        {
            get { return _year; }
            set { _year = value; }
                      
        }
        public int IdStyle
        {
            get { return _idStyle; }
            set { _idStyle = value; }
        }
        public List<Paintings> ListPaintings
        {
            get { return _listPaintings; }
            set { _listPaintings = value; }
        }

        public Paintings() { }

        public List<Paintings> UploadPaintingsFromExcel()
        {
            _listPaintings = new List<Paintings>();
            Workbook wb;
            try
            {
                wb = new Workbook("LR5-var11.xls");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

            WorksheetCollection wbCollection = wb.Worksheets;

            for (int k = 0; k < wbCollection.Count - 2; ++k)
            {
                Worksheet sheet = wbCollection[k];
                int rows = sheet.Cells.MaxDataRow;

                for (int i = 1; i <= rows; ++i)
                {
                    if (sheet.Cells[i, 0].Value == null || string.IsNullOrEmpty(sheet.Cells[i, 0].StringValue))
                    {
                        continue; 
                    }

                    Paintings painting = new Paintings();

                    try
                    {
                        painting.ID = sheet.Cells[i, 0].IntValue;
                        painting.NamePainting = sheet.Cells[i, 1].StringValue;
                        painting.IdArtist = sheet.Cells[i, 2].IntValue;
                        painting.PartOdHermitage = sheet.Cells[i, 3].IntValue;

                    
                        if (sheet.Cells[i, 4].Value != null && !string.IsNullOrEmpty(sheet.Cells[i, 4].StringValue))
                        {
                            string yearValue = sheet.Cells[i, 4].StringValue.Trim();

                            // Пытаемся распарсить как число
                            if (int.TryParse(yearValue, out int year))
                            {
                                painting.Year = year;
                            }
                            else
                            {
                                // Пытаемся извлечь год из текста типа "18 век"
                                painting.Year = ExtractYearFromText(yearValue);
                                if (painting.Year == 0)
                                {
                                    Console.WriteLine($"Неверный формат года в строке {i}: '{yearValue}'");
                                }
                            }
                        }

                        
                        if (sheet.Cells[i, 5].Value != null && !string.IsNullOrEmpty(sheet.Cells[i, 5].StringValue))
                        {
                            string styleValue = sheet.Cells[i, 5].StringValue;
                            if (int.TryParse(styleValue, out int styleId))
                            {
                                painting.IdStyle = styleId;
                            }
                        }

                        _listPaintings.Add(painting);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при обработке строки {i}: {ex.Message}");
                        continue;
                    }
                }
            }
            Console.WriteLine("Успешно прочитаны данные из листа Картины!");
            return _listPaintings;
        }

        // Вспомогательный метод для извлечения года из текста
        private int ExtractYearFromText(string text)
        {
            text = text.ToLower().Trim();
            
            if (text.Contains("век"))
            {
                if (text.Contains("18"))
                    return 1750;
                if (text.Contains("19"))
                    return 1850; 
                if (text.Contains("17"))
                    return 1650; 
                if (text.Contains("16"))
                    return 1550;
            }

            return 0;
        }



        public override string ToString()
        {
            return "ID картины: " + ID + "\nНазвание и год картины: " + NamePainting + " " + Year;
        }


    }
}
