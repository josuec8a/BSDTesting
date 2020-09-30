using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BSD_DataProcessing
{
    public class Constants
    {
        public static string ApiUrl => "https://www.regulations.gov/contentStreamer";

        public static List<Fields> GetFields = new List<Fields>(){
            new Fields("HTS #", "O6", "text"),
            new Fields("Steel Class", "I6", "text"),
            new Fields("Requestor", "F8", "text"),
            new Fields("Importer", "N8", "text"),
            new Fields("Req Qty (kg)", "N28", "number"),
            new Fields("Comments (1d)", "F29", "text"),
            new Fields("Comments (2b)", "D34", "text"),
            new Fields("Product Desc", "C48", "text"),
            new Fields("Aluminum Min", "D67", "%"),
            new Fields("Aluminum Max", "D68", "%"),
            new Fields("Antimony Min", "E67", "%"),
            new Fields("Antimony Max", "E68", "%"),
            new Fields("Bismuth Min", "F67", "%"),
            new Fields("Bismuth Max", "F68", "%"),
            new Fields("Boron Min", "G67", "%"),
            new Fields("Boron Max", "G68", "%"),
            new Fields("Carbon Min", "H67", "%"),
            new Fields("Carbon Max", "H68", "%"),
            new Fields("Chromium Min", "I67", "%"),
            new Fields("Chromium Max", "I68", "%"),
            new Fields("Cobalt Min", "J67", "%"),
            new Fields("Cobalt Max", "J68", "%"),
            new Fields("Copper Min", "K67", "%"),
            new Fields("Copper Max", "K68", "%"),
            new Fields("Iron Min", "L67", "%"),
            new Fields("Iron Max", "L68", "%"),
            new Fields("Lead Min", "M67", "%"),
            new Fields("Lead Max", "M68", "%"),
            new Fields("Manganese Min", "N67", "%"),
            new Fields("Manganese Max", "N68", "%"),
            new Fields("Moly Min", "O67", "%"),
            new Fields("Moly Max", "O68", "%"),
            new Fields("Nickel Min", "P67", "%"),
            new Fields("Nickel Max", "P68", "%"),
            new Fields("Niobium Min", "D70", "%"),
            new Fields("Niobium Max", "D71", "%"),
            new Fields("Nitrogen Min", "E70", "%"),
            new Fields("Nitrogen Max", "E71", "%"),
            new Fields("Phosphorous Min", "F70", "%"),
            new Fields("Phosphorous Max", "F71", "%"),
            new Fields("Selenium Min", "G70", "%"),
            new Fields("Selenium Max", "G71", "%"),
            new Fields("Silicon Min", "H70", "%"),
            new Fields("Silcon Max", "H71", "%"),
            new Fields("Sulfur Min", "I70", "%"),
            new Fields("Sulfur Max", "I71", "%"),
            new Fields("Tellurium Min", "J70", "%"),
            new Fields("Tellurium Max", "J71", "%"),
            new Fields("Titanium Min", "K70", "%"),
            new Fields("Titanium Max", "K71", "%"),
            new Fields("Tungsten Min", "L70", "%"),
            new Fields("Tungsten Max", "L71", "%"),
            new Fields("Vanadium Min", "M70", "%"),
            new Fields("Vanadium Max", "M71", "%"),
            new Fields("Zirconium Min", "N70", "%"),
            new Fields("Zirconium Max", "N71", "%"),
            new Fields("Wall Thickness mm (min)", "D77", "number"),
            new Fields("Wall Thickness mm (max)", "D78", "number"),
            new Fields("ID mm (min)", "E77", "number"),
            new Fields("ID mm (max)", "E78", "number"),
            new Fields("OD mm (min)", "F77", "number"),
            new Fields("OD mm (max)", "F78", "number"),
            new Fields("Length (min)", "G77", "number"),
            new Fields("Length (max)", "G78", "number"),
            new Fields("Tensile Strength in MPa (min)", "J77", "number"),
            new Fields("Tensile Strength in MPa (max)", "J78", "number"),
            new Fields("Yield Strength in MPa (min)", "K77", "number"),
            new Fields("Yield Strength in MPa (max)", "K78", "number"),
            new Fields("Elongation % (min)", "D83", "%"),
            new Fields("Comment (4b)", "C95", "text"),
            new Fields("Comment (4c)", "C97", "text"),
            new Fields("Country of Origin 1", "D100", "text"),
            new Fields("Country of Origin 2", "D101", "text"),
            new Fields("Country of Origin 3", "D102", "text"),
            new Fields("Country of Origin 4", "D103", "text"),
            new Fields("Country of Origin 5", "D104", "text"),
            new Fields("Current Manufacturer 1", "K100", "text"),
            new Fields("Current Manufacturer 2", "K101", "text"),
            new Fields("Current Manufacturer 3", "K102", "text"),
            new Fields("Current Manufacturer 4", "K103", "text"),
            new Fields("Current Manufacturer 5", "K104", "text"),
            new Fields("Country of Export 1", "G100", "text"),
            new Fields("Country of Export 2", "G101", "text"),
            new Fields("Country of Export 3", "G102", "text"),
            new Fields("Country of Export 4", "G103", "text"),
            new Fields("Country of Export 5", "G104", "text"),
            new Fields("Exclusion Quantity 1", "I100", "text"),
            new Fields("Exclusion Quantity 2", "I101", "text"),
            new Fields("Exclusion Quantity 3", "I102", "text"),
            new Fields("Exclusion Quantity 4", "I103", "text"),
            new Fields("Exclusion Quantity 5", "I104", "text")};
    }

    public class Fields
    {
        public Fields(string name, string cellPosition, string formatType)
        {
            Name = name;
            CellPosition = cellPosition;
            FormatType = formatType;
        }
        public string Name { get; set; }
        public string CellPosition { get; set; }
        public string FormatType { get; set; }
    }
}
