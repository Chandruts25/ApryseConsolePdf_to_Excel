using System;
using System.Data;
using System.Diagnostics;
using ApryseConsolePdf_to_Excel;
using pdftron;
using pdftron.Common;
using pdftron.PDF;
using Dapper;
using System.Data.SqlClient;

namespace PDF2Office
{
    // The Apryse SDK Structured Output add-on module can be downloaded from
    // https://docs.apryse.com/documentation/core/info/modules/

    class PdfToExcel
    {

        static PDFNetLoader pdfNetLoader = PDFNetLoader.Instance();
        static Response response = new();

        const string inputPath = @"C:\Users\hp\Desktop\ExclusionPDFComplete\PDF_SourceFolder\";
        const string outputPath = @"C:\Users\hp\Desktop\ExclusionPDFComplete\EXCEL_ExtractedFolder\";

        static int Main(string[] args)
        {
            try
            {
                PDFNet.Initialize("demo:1721917547091:7e6e33d60300000000a1be6388807171fb0a9afe315df0f72c6f73122d");
                PDFNet.AddResourceSearchPath(@"C:\Users\hp\Desktop\StructuredOutputWindows\Lib\");

                if (!StructuredOutputModule.IsModuleAvailable())
                {
                    Console.WriteLine("Apryse SDK Structured Output module not available.");
                    return 0;
                }

                //    List<string> fileNames = new List<string>
                //{
                //    "2024_07_01_-Wyoming-Medicaid-Exclusion-List-July",
                //    "AlaskaExcludedProviderList",
                //    "Idaho Medicaid Exclusion List",
                //    "Med Prov Excl-Rein List-UPDATED-07.10.2024",
                //    "Medicaid Excluded Providers",
                //    "nj_debarment_list (1)",
                //    "provider-exclusion-list",
                //    "ProviderSuspensionsTerminations",
                //    "terminatedproviderlist",
                //    "WV Medicaid Provider Exclusions and Terminations July 2024"
                //};

                string[] filePaths = Directory.GetFiles(inputPath, "*.pdf");

                if (filePaths.Length == 0)
                {
                    Console.WriteLine("No files found in the specified directory.");
                    return 0;
                }
                else
                {
                    foreach (var filePath in filePaths)
                    {
                        var fileName = Path.GetFileNameWithoutExtension(filePath);

                        string outputFile = outputPath + fileName + ".xlsx";

                        pdftron.PDF.Convert.ExcelOutputOptions options = new pdftron.PDF.Convert.ExcelOutputOptions();
                        //options.SetNonTableContent(true);
                        options.SetSingleSheet(true);
                        //options.SetPageSingleSheet(true);
                        //options.SetHeadersAndFootersSetting(pdftron.PDF.Convert.StructuredOutputOptions.SectionConversionSetting.e_Recover);

                        pdftron.PDF.Convert.ToExcel(filePath, outputFile, options);
                        Console.WriteLine("File Name : " + fileName + ".pdf processed.");

                        response = new Response
                        {
                            FileName = fileName,
                            InputUrl = filePath,
                            OutputUrl = outputFile,
                            IsSuccess = true,
                        };
                        AddResponse(response);
                    }
                }
            }
            catch (PDFNetException e)
            {
                Console.WriteLine("Unable to convert PDF document to Excel, error: " + e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unknown Exception, error: " + e.Message);
            }
            finally
            {
                PDFNet.Terminate();
                Console.WriteLine("Done.");
            }
            return 0;
        }


        static void AddResponse(Response response)
        {
            var parameters = new DynamicParameters();
            parameters.Add("@FileName", response.FileName);
            parameters.Add("@InputUrl", response.InputUrl);
            parameters.Add("@OutputUrl", response.OutputUrl);
            parameters.Add("@ErrorMessage", response.ErrorMessage);
            parameters.Add("@IsSuccess", response.IsSuccess);

            using var con = new SqlConnection("Server=LAPTOP-5SETUAID\\SQLEXPRESS;Database=PdfTtoExcel;Trusted_Connection=True;TrustServerCertificate=yes");

            con.Open();
            con.Execute("InsertResponse", parameters, commandType: CommandType.StoredProcedure);
            con.Close();

            Console.WriteLine("Data inserted successfully.");
        }

    }
}
