using System.Collections.Generic;

namespace Colugo.Models
{
    public class ReadmeModel
    {
        public string DocumentTitle => "LEED v4.1 Building Design and Construction (BD+C)";
        public string Version => "v4.1 Beta - January 2019";
        public string IssuedBy => "U.S. Green Building Council (USGBC)";
        public int TotalPoints => 110;

        public string GetOverview()
        {
            return
                "LEED v4.1 BD+C (Building Design and Construction) is a rating system developed by USGBC\n" +
                "for high-performance green buildings. It evaluates buildings across multiple environmental\n" +
                "and sustainability categories, with 110 total points available.\n\n" +
                "Four key goals guide LEED v4.1:\n" +
                "  1. Ensure Leadership\n" +
                "  2. Increase Achievability\n" +
                "  3. Measure Performance\n" +
                "  4. Expand the Market\n\n" +
                "The rating system applies to: New Construction, Core and Shell, Schools, Retail,\n" +
                "Data Centers, Warehouses and Distribution Centers, Hospitality, and Healthcare.";
        }

        public string GetCertificationLevels()
        {
            return
                "LEED Certification Levels (Total 110 Points):\n" +
                "  - Certified:  40 - 49 points\n" +
                "  - Silver:     50 - 59 points\n" +
                "  - Gold:       60 - 79 points\n" +
                "  - Platinum:   80+ points";
        }

        public List<ScoreCategory> GetScorecard()
        {
            return new List<ScoreCategory>
            {
                new ScoreCategory("IP", "Integrative Process", 1, 1,
                    "Integrative Project Planning and Design (P), Integrative Process (1 pt)"),

                new ScoreCategory("LT", "Location and Transportation", 16, 1,
                    "LEED for Neighborhood Development Location (16 pts), Sensitive Land Protection (1 pt), " +
                    "High-Priority Site (2 pts), Surrounding Density and Diverse Uses (5 pts), " +
                    "Access to Quality Transit (5 pts), Bicycle Facilities (1 pt), " +
                    "Reduced Parking Footprint (1 pt), Electric Vehicles (1 pt)"),

                new ScoreCategory("SS", "Sustainable Sites", 10, 2,
                    "Construction Activity Pollution Prevention (P), Site Assessment (1 pt), " +
                    "Protect or Restore Habitat (2 pts), Open Space (1 pt), " +
                    "Rainwater Management (3 pts), Heat Island Reduction (2 pts), " +
                    "Light Pollution Reduction (1 pt)"),

                new ScoreCategory("WE", "Water Efficiency", 11, 3,
                    "Outdoor Water Use Reduction (P + 2 pts), Indoor Water Use Reduction (P + 6 pts), " +
                    "Building-Level Water Metering (P), Cooling Tower Water Use (2 pts), " +
                    "Water Metering (1 pt)"),

                new ScoreCategory("EA", "Energy and Atmosphere", 33, 4,
                    "Fundamental Commissioning and Verification (P), Minimum Energy Performance (P), " +
                    "Building-Level Energy Metering (P), Fundamental Refrigerant Management (P), " +
                    "Enhanced Commissioning (6 pts), Optimize Energy Performance (18 pts), " +
                    "Advanced Energy Metering (1 pt), Grid Harmonization (2 pts), " +
                    "Renewable Energy (5 pts), Enhanced Refrigerant Management (1 pt)"),

                new ScoreCategory("MR", "Materials and Resources", 13, 2,
                    "Storage and Collection of Recyclables (P), Construction and Demolition Waste Management Planning (P), " +
                    "Building Life-Cycle Impact Reduction (5 pts), " +
                    "BPDO - EPD (2 pts), BPDO - Sourcing of Raw Materials (2 pts), " +
                    "BPDO - Material Ingredients (2 pts), " +
                    "Construction and Demolition Waste Management (2 pts)"),

                new ScoreCategory("EQ", "Indoor Environmental Quality", 16, 2,
                    "Minimum Indoor Air Quality Performance (P), Environmental Tobacco Smoke Control (P), " +
                    "Enhanced Indoor Air Quality Strategies (2 pts), Low-Emitting Materials (3 pts), " +
                    "Construction Indoor Air Quality Management Plan (1 pt), " +
                    "Indoor Air Quality Assessment (2 pts), Thermal Comfort (1 pt), " +
                    "Interior Lighting (2 pts), Daylight (3 pts), Quality Views (1 pt), " +
                    "Acoustic Performance (1 pt)"),

                new ScoreCategory("IN", "Innovation", 6, 0,
                    "Innovation (5 pts), LEED Accredited Professional (1 pt)"),

                new ScoreCategory("RP", "Regional Priority", 4, 0,
                    "Regional Priority (4 pts)")
            };
        }
    }

    public class ScoreCategory
    {
        public string Abbreviation { get; }
        public string Name { get; }
        public int MaxPoints { get; }
        public int PrerequisiteCount { get; }
        public string Details { get; }

        public ScoreCategory(string abbreviation, string name, int maxPoints, int prerequisiteCount, string details)
        {
            Abbreviation = abbreviation;
            Name = name;
            MaxPoints = maxPoints;
            PrerequisiteCount = prerequisiteCount;
            Details = details;
        }

        public override string ToString()
        {
            string prereq = PrerequisiteCount > 0 ? $" ({PrerequisiteCount} prerequisites)" : "";
            return $"[{Abbreviation}] {Name}: {MaxPoints} pts{prereq}";
        }
    }
}
