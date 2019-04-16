using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace IOStreams
{

    public static class TestTasks
    {
        /// <summary>
        /// Parses Resourses\Planets.xlsx file and returns the planet data: 
        ///   Jupiter     69911.00
        ///   Saturn      58232.00
        ///   Uranus      25362.00
        ///    ...
        /// See Resourses\Planets.xlsx for details
        /// </summary>
        /// <param name="xlsxFileName">source file name</param>
        /// <returns>sequence of PlanetInfo</returns>
        public static IEnumerable<PlanetInfo> ReadPlanetInfoFromXlsx(string xlsxFileName)
        {
            Package z = ZipPackage.Open(xlsxFileName);
            XNamespace xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            var sharedString = z.GetPart(new Uri("/xl/sharedStrings.xml", UriKind.Relative));
            XDocument xdocSharedString = XDocument.Load(sharedString.GetStream());
            var name =  xdocSharedString.Root.Elements().Select(x => (string)x);


            var worksheet = z.GetPart(new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative));
            XDocument xdocWorksheet = XDocument.Load(worksheet.GetStream());
            var value = xdocWorksheet.Root.Element(xmlns + "sheetData").Elements(xmlns + "row").
                Where(x => Int32.Parse(x.Attribute("r").Value) >= 2).SelectMany(x => x.Elements()).
                Where(x => ((string)x.Attribute("r"))[0] == 'B').Select(x=>(double)x.Element(xmlns+"v"));


            z.Close();
            return name.Zip(value, (x, y) => new PlanetInfo() { Name = x, MeanRadius = y });


            // TODO : Implement ReadPlanetInfoFromXlsx method using System.IO.Packaging + Linq-2-Xml

            // HINT : Please be as simple & clear as possible.
            //        No complex and common use cases, just this specified file.
            //        Required-требуемее data are stored-хранятся in Planets.xlsx archive in 2 files:
            //         /xl/sharedStrings.xml      - dictionary of all string values
            //         /xl/worksheets/sheet1.xml  - main worksheet
            // throw new NotImplementedException();
        }


        /// <summary>
        /// Calculates hash of stream using specifued algorithm
        /// </summary>
        /// <param name="stream">source stream</param>
        /// <param name="hashAlgorithmName">hash algorithm ("MD5","SHA1","SHA256" and other supported by .NET)</param>
        /// <returns></returns>
        public static string CalculateHash(this Stream stream, string hashAlgorithmName)
        {
            var hashAlgoritm = HashAlgorithm.Create(hashAlgorithmName);
            if (hashAlgoritm == null)
                throw new ArgumentException();
            
            var hash = hashAlgoritm.ComputeHash(stream);

            return BitConverter.ToString(hash).Replace("-", String.Empty);

            // TODO : Implement CalculateHash method
            throw new NotImplementedException();
        }


        /// <summary>
        /// Returns decompressed strem from file. 
        /// </summary>
        /// <param name="fileName">source file</param>
        /// <param name="method">method used for compression (none, deflate, gzip)</param>
        /// <returns>output stream</returns>
        public static Stream DecompressStream(string fileName, DecompressionMethods method)
        {
            Func<Stream, DecompressionMethods, Stream> GetDecompressStream = (x, decompressionMethod) =>
             {
                 switch (decompressionMethod)
                 {
                     case DecompressionMethods.Deflate:
                         return new DeflateStream(x, CompressionMode.Decompress);
                     case DecompressionMethods.GZip:
                         return new GZipStream(x, CompressionMode.Decompress);
                     case DecompressionMethods.None:
                         return x;
                     default:
                         throw new ArgumentException("check decompression method!");
                 }
             };


            FileStream sourceStream = new FileStream(fileName, FileMode.OpenOrCreate);         
            return  GetDecompressStream(sourceStream, method);         
            
            // TODO : Implement DecompressStream method
            //throw new NotImplementedException();
        }



        /// <summary>
        /// Reads file content econded with non Unicode encoding
        /// </summary>
        /// <param name="fileName">source file name</param>
        /// <param name="encoding">encoding name</param>
        /// <returns>Unicoded file content</returns>
        public static string ReadEncodedText(string fileName, string encoding)
        {
            return File.ReadAllText(fileName, Encoding.GetEncoding(encoding));

            // TODO : Implement ReadEncodedText method
            //throw new NotImplementedException();
        }
    }


    public class PlanetInfo : IEquatable<PlanetInfo>
    {
        public string Name { get; set; }
        public double MeanRadius { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1}", Name, MeanRadius);
        }

        public bool Equals(PlanetInfo other)
        {
            return Name.Equals(other.Name)
                && MeanRadius.Equals(other.MeanRadius);
        }
    }



}
