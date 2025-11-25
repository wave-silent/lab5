using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

namespace lab5
{
    internal class Artists
    {
        private int _id;
        private string _name;
        private List<Artists> _listArtists;
        
        public int ID
        {
            get { return _id; }
            set { _id = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }  
        }

        public List<Artists> ListArtists
        {
            get { return _listArtists; }
            set { _listArtists = value; }
        }

        public Artists() { }
        
        public Artists(int id, string name)
        {
            ID = id;
            Name = name;      
        }

        public List<Artists> UploadArtistsFromExcel()
        {
            _listArtists = new List<Artists>();
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

            for (int k = 1; k < wbCollection.Count-1; ++k)
            {
                Worksheet sheet = wbCollection[k];
                int rows = sheet.Cells.MaxDataRow;       

                for (int i = 1; i <= rows; ++i)
                {                
                    _listArtists.Add(new Artists(sheet.Cells[i, 0].IntValue, sheet.Cells[i, 1].StringValue));                   
                }
            }
            Console.WriteLine("Успешно прочитаны данные из листа Художники!");
            return _listArtists;
        }

        public override string ToString()
        {
            return "ID мастера: " + ID + "\nИмя и фамилия мастера: " + Name;
        }
    }
}
