using ExcelProject.Entities;
using System.Collections.Generic;

namespace ExcelProject.DataAccess.Abstract
{
    public interface IFileDal
    {        List<T_GK_MAIN> Read(string filePath);
    }
}
