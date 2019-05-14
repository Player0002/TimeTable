
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 시간표
{
    static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AppDomain.CurrentDomain.AssemblyResolve += Assembly_Resolve;

            LoadDll("MaterialSkin.dll"); // 로드할 dll 이름을 적는곳 ("ABC.dll") 이런식으로
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
        static Dictionary<string, Assembly> dict = new Dictionary<string, Assembly>();

        static bool LoadDll(string path)

        {

            Assembly curAssm = Assembly.GetExecutingAssembly();

            string appName = curAssm.GetName().Name.Replace(" ", "_");

            Assembly dllAssm = null;

            byte[] dllData;

            using (System.IO.Stream s = curAssm.GetManifestResourceStream($"{appName}.{path}"))

            {

                if (s != null)

                {

                    dllData = new byte[s.Length];

                    s.Read(dllData, 0, (int)s.Length);

                    dllAssm = Assembly.Load(dllData);

                }

                else

                {

                    return false;

                }

            }

            dict.Add(dllAssm.FullName, dllAssm);

            return true;

        }

        static Assembly Assembly_Resolve(object sender, ResolveEventArgs e)

        {

            if (dict.ContainsKey(e.Name))

            {

                Assembly assm = dict[e.Name];

                dict.Remove(e.Name);

                return assm;

            }

            return null;

        }
    }
}