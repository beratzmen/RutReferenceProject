using ExcelProject.DataAccess.Abstract;
using ExcelProject.DataAccess.Concrete.Context;
using ExcelProject.Entities;
using System.Collections.Generic;

namespace ExcelProject.DataAccess.Concrete.Repository
{
    public class FileRepository : IFileDal
    {
        Excel context = new Excel();
        public List<T_GK_MAIN> Read(string filePath)
        {
            return context.Read(filePath);
        }
    }
}
