using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Main
{
    class Class3
    {

        static void Main(string[] args)
        {
            //Dentist_Scraper.Program prog = new Dentist_Scraper.Program();
            Aetna_Scraper.AetnaProgram prog = new Aetna_Scraper.AetnaProgram();
            prog.DentistParser();
        }


    }
}
