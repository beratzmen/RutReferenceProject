using ExcelProject.DataAccess.Abstract;
using ExcelProject.Entities;
using System.Collections.Generic;
using System.Linq;

namespace ExcelProject.Business
{
    public class FileManager
    {
        IFileDal _fileDal;
        public FileManager(IFileDal fileDal)
        {
            _fileDal = fileDal;
        }

        /// Kaynak yolu verilen dosyanın içeriği okur.       
        public List<T_GK_MAIN> Read(string filePath)
        {
            List<T_GK_MAIN> mainItems = _fileDal.Read(filePath);
            if (mainItems == null)
                return null;
            mainItems.ForEach(p => PopulateTGK(p));
            return mainItems;
        }

        /// Gelen modelin tekil grup kodlarını birbiriyle çarpıp Entity tipindeki listeye doldurur.
        public void PopulateTGK(T_GK_MAIN mainItems)
        {
            var emptyData = new T_GK() { code = " " };
            int i = 3;
            mainItems.allItems = (from tgk1 in mainItems.tgk1.DefaultIfEmpty(emptyData)
                                  from tgk2 in mainItems.tgk2.DefaultIfEmpty(emptyData)
                                  from tgk3 in mainItems.tgk3.DefaultIfEmpty(emptyData)
                                  from tgk4 in mainItems.tgk4.DefaultIfEmpty(emptyData)
                                  from tgk5 in mainItems.tgk5.DefaultIfEmpty(emptyData)
                                  select new Entity
                                  {
                                      id = i++,
                                      code1 = tgk1.code,
                                      code2 = tgk2.code,
                                      code3 = tgk3.code,
                                      code4 = tgk4.code,
                                      code5 = tgk5.code
                                  }).ToList();
        }
    }
}
