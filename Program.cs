using System.Collections.Generic;
using System.Linq;
using System;
using ClosedXML.Excel;

namespace match
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new XLWorkbook("../excel/pair.xlsx");
            var joi_sheet = workbook.Worksheet("上位");
            var kai_sheet = workbook.Worksheet("下位");
            var joi_students = get_students(joi_sheet);
            var kai_students = get_students(kai_sheet);
            var pair_dict = new Dictionary<string, string>();

            foreach (var st in joi_students)
            {
                foreach (var pair in st.hope_list)
                {
                    if (pair_dict.ContainsKey(st.name))
                    {
                        break;
                    }

                    // check pair 既にペアができているか？
                    if (pair_dict.ContainsValue(pair))
                    {
                        // 下位の生徒の中での、　上位の生徒のランクを比較
                        var kai_student = kai_students.Where(x => x.name == pair).FirstOrDefault();
                        var joi_student = pair_dict.FirstOrDefault(x => x.Value.Equals(pair)).Key;

                        if (kai_student.hope_list.IndexOf(st.name) <= (kai_student.hope_list.IndexOf(joi_student)))
                        {
                            pair_dict.Remove(joi_student);
                            pair_dict.Add(st.name, pair);
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    pair_dict.Add(st.name, pair);
                }
            }

            Console.WriteLine("===pair===");

            foreach (var item in pair_dict)
            {
                Console.WriteLine($"{item.Key} {item.Value}");
            }
        }

        public class Student
        {
            public string name {get; set;}
            public List<string> hope_list { get; set;}

            public Student()
            {
                hope_list = new List<string>();
            }
        }

        public static List<Student> get_students(ClosedXML.Excel.IXLWorksheet ws)
        {
            var students = new List<Student>();

            for (int row = 3; row <= ws.CellsUsed().Count(); row++)
            {
                var name = ws.Cell(row, 1).Value.ToString();

                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                var st = new Student();
                st.name = name.Trim();

                for (int col = 2; col <= ws.CellsUsed().Count(); col++)
                {
                    var pair = ws.Cell(row, col).Value.ToString();

                    if (string.IsNullOrEmpty(pair))
                    {
                        continue;
                    }
                    st.hope_list.Add(pair.Trim());
                }

                students.Add(st);
            }

            return students;
        }

    }
}
