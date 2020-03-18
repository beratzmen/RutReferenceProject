using System.Collections.Generic;

namespace ExcelProject.Entities
{
    public class T_GK_MAIN
    {
        public int id { get; set; }
        public List<Entity> allItems { get; set; }
        public List<T_GK> tgk1 { get; set; }
        public List<T_GK> tgk2 { get; set; }
        public List<T_GK> tgk3 { get; set; }
        public List<T_GK> tgk4 { get; set; }
        public List<T_GK> tgk5 { get; set; }
    }
    public class T_GK
    {
        public string code { get; set; }
    }
    public class Entity
    {
        public int id { get; set; }
        public string code1 { get; set; }
        public string code2 { get; set; }
        public string code3 { get; set; }
        public string code4 { get; set; }
        public string code5 { get; set; }
    }
}
