using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using CommandLine;


namespace ExcelBridge
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new Options();
            var isValid = CommandLine.Parser.Default.ParseArguments(args, options);
            if(isValid)
            {
                ExcelAPI api = new ExcelAPI(options.InputFile);
                switch (options.command)
                {
                    //Scrive il valore su una cella
                    case "writetocell":
                        api.writeToCell(options.row, options.col, options.value);
                        break;
                    //Esegue una macro
                    case "runmacro":
                        api.runMacro(options.macro);
                        break;
                    //Legge una cella e stampa a video il risultato
                    case "readcell":
                        Console.Write(api.getCellValue(options.row, options.col).ToString());
                        break;
                }
                if (options.calculate)
                    api.calculate();
                api.save();
                api.close();
            }
        }        
    }

    //Classe che rappresenta i parametri della CLI(Command Line Interface)
    class Options
    {
        //Il file di input
        [Option("file", Required = true, HelpText = "Input .xlsx file")]
        public string InputFile { get; set; }

        //Il comando da eseguire
        [Option("command", Required = true, HelpText = "Command")]
        public string command { get; set; }

        //Il foglio su cui eseguire il comando, default 1
        [Option("sheet", Required = false, HelpText = "The active sheet", DefaultValue = 1)]
        public int sheet { get; set; }

        //Abilita il calcolo alla fine del comando
        [Option("calculate", Required = false, HelpText = "Enable calculation at the end of command", DefaultValue = false)]
        public Boolean calculate { get; set; }

        //La riga
        [Option("row", Required = false, HelpText = "The row")]
        public int row { get; set; }

        //La Colonna
        [Option("col", Required = false, HelpText = "The col")]
        public int col { get; set; }

        //Il valore
        [Option("value", Required = false, HelpText = "Value")]
        public string value { get; set; }

        //Il nome della macro
        [Option("macro", Required = false, HelpText = "Macro name")]
        public string macro { get; set; }
    }
}
