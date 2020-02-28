using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelBridge
{
    class ExcelAPI
    {
        private String filename;
        private Application excel;
        private Workbook workBook;
        private Worksheet sheet;

        /**
         * Costruisce l'oggetto Excel API
         * Apre il file passato come parametro, seleziona il primo foglio
         * e disabilita di default il calcolo automatico
         **/
        public ExcelAPI(string filename)
        {
            this.filename = filename;
            this.excel = new Application();
            this.workBook = excel.Workbooks.Open(filename);
            this.changeSheet(1);
            this.disableAutomaticCalculation();
        }

        /**
         * Cambia il foglio di default
         */
        public void changeSheet(int sheetNumber)
        {
            this.sheet = (Worksheet)this.workBook.Sheets[1];
        }

        /**
         * Scrive un valore su una cella
         */
        public void writeToCell(int i, int j, Object value)
        {
            sheet.Cells[i, j].Value = value;
        }

        /**
         *  Legge il valore da una cella
         */
        public Object getCellValue(int i, int j)
        {
            return sheet.Cells[i, j].Value;
        }


        /**
         * Esegue una Macro 
         */
        public void runMacro(String macroname)
        {
            excel.Run(macroname);
        }

        /**
         * Salva le modifiche nel file corrente 
         */
        public void save()
        {
            workBook.Save();
        }

        /**
         * Chiude il documento corrente 
         * 
         */
        public void close()
        {
            this.workBook.Close();
        }

        /**
         * Disabilita il calcolo automatico
         * 
         */
        public void disableAutomaticCalculation()
        {
            this.excel.Calculation = XlCalculation.xlCalculationManual;
        }

        /**
         * Abilita il calcolo automatico
         * 
         */
        public void enableAutomaticCalculation()
        {
            this.excel.Calculation = XlCalculation.xlCalculationAutomatic;
        }


        /**
         * Esegue il calcolo 
         */
        public void calculate()
        {
            this.excel.Calculate();
        }
    }
}
