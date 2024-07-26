using System;
using pdftron;
using pdftron.Common;
using pdftron.PDF;

namespace PDF2OfficeTestCS
{
    // The Apryse SDK Structured Output add-on module can be downloaded from
    // https://docs.apryse.com/documentation/core/info/modules/

    class Class1
    {

        private static PDFNetLoader pdfNetLoader = PDFNetLoader.Instance();

        static Class1() { }

        //const string inputPath = @"C:\Users\hp\Desktop\ExclusionPDFComplete\ExclusionPDFComplete\";
        const string inputPath = @"C:\Users\hp\Downloads\";
        const string outputPath = inputPath;

        //[STAThread]
        static int Main(string[] args)
        {
            PDFNet.Initialize("demo:1721917547091:7e6e33d60300000000a1be6388807171fb0a9afe315df0f72c6f73122d");

            PDFNet.AddResourceSearchPath(@"C:\Users\hp\Desktop\StructuredOutputWindows\Lib\");

            if (!StructuredOutputModule.IsModuleAvailable())
            {
                Console.WriteLine("Apryse SDK Structured Output module not available.");
                return 0;
            }

            bool err = false;

            // Excel
            List<string> fileNames = new List<string>
            {
                "2024_07_01_-Wyoming-Medicaid-Exclusion-List-July",
                "AlaskaExcludedProviderList",
                "Idaho Medicaid Exclusion List",
                "Med Prov Excl-Rein List-UPDATED-07.10.2024",
                "Medicaid Excluded Providers",
                "nj_debarment_list (1)",
                "provider-exclusion-list",
                "ProviderSuspensionsTerminations",
                "terminatedproviderlist",
                "WV Medicaid Provider Exclusions and Terminations July 2024"
            };

            //foreach (var fileName in fileNames)
            //{
            try
            {
                //var fileName = "terminatedproviderlist";
                var fileName = "6546565464";

                string outputFile = outputPath + fileName + ".xlsx";

                pdftron.PDF.Convert.ExcelOutputOptions options = new pdftron.PDF.Convert.ExcelOutputOptions();

                options.SetNonTableContent(true);
                // For Add header for all tables
                options.SetSingleSheet(true);
                // For multiple sheets
                options.SetPageSingleSheet(true);
                options.SetHeadersAndFootersSetting(pdftron.PDF.Convert.StructuredOutputOptions.SectionConversionSetting.e_Recover);

                pdftron.PDF.Convert.ToExcel(inputPath + fileName + ".pdf", outputFile, options);

                Console.WriteLine("Result saved in " + outputFile);
            }
            catch (PDFNetException e)
            {
                Console.WriteLine("Unable to convert PDF document to Excel, error: " + e.Message);
                err = true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Unknown Exception, error: ");
                Console.WriteLine(e);
                err = true;
            }
            //}

            PDFNet.Terminate();
            Console.WriteLine("Done.");
            return (err == false ? 0 : 1);
        }
    }
}
