using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace ImportToDatabase.Models
{
    public class ProductCatalog
    {
        public int Id { get; set; }
        public int? CustomerId { get; set; }
        public string Description { get; set; }
        public string NMFC { get; set; }
        public bool IsHazardous { get; set; }
        public decimal? Class { get; set; }
        public decimal? Length { get; set; }
        public decimal? Width { get; set; }
        public decimal? Height { get; set; }
        public decimal? Density { get; set; }
        public string DimensionUnit { get; set; }
        public decimal? Weight { get; set; }
        public string WeightUnit{ get; set; }
        public string HazmatContact { get; set; }
        public string HazardUnNumber { get; set; }
        public string PackageType { get; set; }
        public string UnitDensity { get; set; }
        // Add other properties as needed
    }
    public enum UnitDensityType
    {
        PCF = 0,
        KGM = 1
    }
    public enum UnitDimentionType
    {
        [Description("Inches")]
        Inches = 0,
        [Description("Feet")]
        Feet = 1,
        [Description("Meters")]
        Meters = 2,
        [Description("Yards")]
        Yards = 3,
        [Description("Centimeters")]
        Centimeters = 4
    }
    public enum UnitWeightType
    {
        [Description("LB")]
        LB = 0,
        [Description("KG")]
        KG = 1,
        [Description("TONS")]
        TONS = 2,
        [Description("TONNAGE")]
        TONNAGE = 3
    }
}
