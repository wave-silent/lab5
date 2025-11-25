using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5
{
    internal class Styles
    {
        private int _id;
        private string _paintingStyle;
        private List<Styles> _listStyles;

        public int ID
        {
            get { return _id; }
            set { _id = value; }
        }

        public string PaintingStyle
        {
            get { return _paintingStyle; }
            set { _paintingStyle = value; }
        }
        public List<Styles> ListStyle
        {
            get { return _listStyles; }
            set { _listStyles = value; }
        }

        public Styles() { }

        public Styles(int id, string paintingStyle)
        {
            ID = id;
            PaintingStyle = paintingStyle;
        }

        public List<Styles> UploadStylesFromExcel()
        {
            _listStyles = new List<Styles>();
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

            for (int k = 2; k < wbCollection.Count; ++k)
            {
                Worksheet sheet = wbCollection[k];
                int rows = sheet.Cells.MaxDataRow;

                for (int i = 1; i <= rows; ++i)
                {
                    Styles style = new Styles();

                    style.ID = sheet.Cells[i, 0].IntValue;

                    if (sheet.Cells[i, 1].Value == null || string.IsNullOrEmpty(sheet.Cells[i, 1].StringValue))
                    {
                        continue;
                    }

                    style.PaintingStyle = sheet.Cells[i, 1].StringValue;


                    _listStyles.Add(new Styles(style.ID, style.PaintingStyle));
                }
            }
            Console.WriteLine("Успешно прочитаны данные из листа Стиль!");
            return _listStyles;
        }

        public override string ToString()
        {
            return "ID стиля: " + ID + "\nСтиль живописи: " + PaintingStyle;
        }
    }
}
