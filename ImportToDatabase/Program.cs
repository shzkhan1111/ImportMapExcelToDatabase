using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using OfficeOpenXml;
using Microsoft.EntityFrameworkCore;
using ImportToDatabase.Models;
using Microsoft.Data.SqlClient;
using System.Linq;

namespace ImportToDatabase
{

    class Program
    {
        //readonly int client= 1;
        //readonly string customername = "UFPT";

        static void Main(string[] args)
        {
            Console.WriteLine("Starting the Excel to Database Import Process...");
            string filePath = @"C:\Users\shahzaib.ahmed.DESKTOP-EB9K6GQ\Downloads\UFPT Product Guide.xlsx";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var productCatalogList = ReadExcelData(filePath);
            InsertDataToDatabase(productCatalogList);

            //using (var context = new ApplicationDBContext())
            //{
            //    //context.ProductCatalogs.AddRange(productCatalogList);
            //    //context.SaveChanges();
            //    Console.WriteLine("Data uploaded successfully!");
            //}

        }

        static decimal? ParseDecimal(string input)
        {
            return decimal.TryParse(input, out var result) ? result : (decimal?)null;
        }

        static List<ProductCatalog> ReadExcelData(string filePath)
        {
            var products = new List<ProductCatalog>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // First worksheet
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    var product = new ProductCatalog();
                    product.Description = worksheet.Cells[row, 2].Text;
                    product.NMFC = worksheet.Cells[row, 3].Text;
                    //product.IsHazardous = bool.Parse(worksheet.Cells[row, 4].Text ?? "false");
                    product.Class = ParseDecimal(worksheet.Cells[row, 4].Text);
                    product.Length = ParseDecimal(worksheet.Cells[row, 5].Text);
                    product.Width = ParseDecimal(worksheet.Cells[row, 6].Text);
                    product.Height = ParseDecimal(worksheet.Cells[row, 7].Text);
                    product.DimensionUnit = worksheet.Cells[row, 8].Text;
                    product.Weight = ParseDecimal(worksheet.Cells[row, 9].Text);
                    product.WeightUnit = worksheet.Cells[row, 10].Text;
                    product.PackageType = worksheet.Cells[row, 11].Text;
                    product.HazardUnNumber = worksheet.Cells[row, 12].Text;
                    product.Density = ParseDecimal(worksheet.Cells[row, 13].Text);
                    product.UnitDensity = worksheet.Cells[row, 14].Text;

                    products.Add(product);
                }
            }

            return products;
        }
        static int? GetCustomerId(SqlConnection connection, string customername)
        {
            int clientloadid = 1861865;
            string query = $"SELECT CustomerId FROM customer WHERE CustomerId = @CustomerId and clientid = {clientloadid}";
            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@CustomerId", 4093902);
                var result = command.ExecuteScalar();
                return result != DBNull.Value ? (int?)result : null;
            }
        }
        static Dictionary<int, string> GetPackageTypes(SqlConnection connection)
        {
            string query = "SELECT Id, Description from PackageType";
            var packageTypes = new Dictionary<int, string>();
            using (SqlCommand command = new SqlCommand(query, connection))
            {
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int packageId = reader.GetInt32(0); // Id
                        string description = reader.GetString(1); // Description
                        packageTypes[packageId] = description;
                    }
                }
            }
            return packageTypes;
        }
        static void InsertDataToDatabase(List<ProductCatalog> productCatalogList)
        {
            // Define the connection string
            string connectionString = "Server=tcp:brokerwareio20170123123330dbserver.database.windows.net,1433;Initial Catalog=3plbrokerwaredb_staging_2023-03-24T13-15Z;Persist Security Info=False; User ID=azuresa@brokerwareio20170123123330dbserver;Password=3pl$y$3760;MultipleActiveResultSets=True";
            //string customername = "Genetech1";
            string customername = "UFPT";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    int? customerId = GetCustomerId(connection, customername);
                    var packageTypes = GetPackageTypes(connection);

                    foreach (var product in productCatalogList)
                    {

                        int? unitsDensity = null;
                        if (Enum.TryParse<UnitDensityType>(product.UnitDensity.ToString(), out var densityEnum))
                        {
                            unitsDensity = (int)densityEnum;
                        }

                        int? unitsDim = null;
                        if (Enum.TryParse<UnitDimentionType>(product.DimensionUnit.ToString(), out var dimensionEnum))
                        {
                            unitsDim = (int)dimensionEnum;
                        }
                        List<string> WeightUnitLBDList = new List<string> { "LBS" };
                        string wunit = WeightUnitLBDList.Any(x => x.Equals(product.WeightUnit, StringComparison.OrdinalIgnoreCase)) ? "LB" : product.WeightUnit;
                       
                        int unitsWeight = (int)Enum.Parse(typeof(UnitWeightType), wunit);

                        
                        string query = @"
                                        INSERT INTO ProductCatalog (
                                            [CustomerId], [Description], [NMFC], [Class], 
                                            [Length], [Width], [Height], [UnitsDensity], 
                                            [Weight], [UnitsWeight], [PackageTypeId], 
                                            [CreatedDate], [ModifiedDate] , [Density] , [IsHazardous] , [UnitsDim]
                                        )
                                        VALUES (
                                            @CustomerId, @Description, @NMFC, @Class, 
                                            @Length, @Width, @Height, @UnitsDensity, 
                                            @Weight, @UnitsWeight, @PackageTypeId, 
                                            @CreatedDate, @ModifiedDate , 0 , 0 , @UnitsDim
                                        )";
                        if (product.Description.Length > 150)
                        {
                            product.Description = product.Description.Substring(0, 150);
                        }

                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            // Add parameters to prevent SQL injection
                            command.Parameters.AddWithValue("@CustomerId", customerId.Value);
                            command.Parameters.AddWithValue("@Description", product.Description ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@NMFC", product.NMFC ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@Class", product.Class ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@Length", product.Length ?? 0); // Ensure Length is nullable
                            command.Parameters.AddWithValue("@Width", product.Width ?? 0); // Ensure Width is nullable
                            command.Parameters.AddWithValue("@Height", product.Height ?? 0); // Ensure Height is nullable
                            command.Parameters.AddWithValue("@UnitsDensity", unitsDensity ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@Weight", product.Weight ?? 0); // Ensure Weight is nullable
                            command.Parameters.AddWithValue("@UnitsWeight", unitsWeight);
                            int? packageTypeId = null;

                            if (packageTypes.Any(x => x.Value.Equals(product.PackageType, StringComparison.OrdinalIgnoreCase)))
                            {
                                packageTypeId = packageTypes.FirstOrDefault(x => x.Value.Equals(product.PackageType, StringComparison.OrdinalIgnoreCase)).Key;
                            }


                            command.Parameters.AddWithValue("@PackageTypeId", packageTypeId ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@UnitsDim", unitsDim);


                            // Set CreatedDate and ModifiedDate to today
                            DateTime currentDate = DateTime.UtcNow;
                            //command.Parameters.AddWithValue("@CreatedBy", "System"); // Or the actual user
                            command.Parameters.AddWithValue("@CreatedDate", currentDate);
                            //command.Parameters.AddWithValue("@ModifiedBy", "System"); // Or the actual user
                            command.Parameters.AddWithValue("@ModifiedDate", currentDate);

                            // Execute the query
                            command.ExecuteNonQuery();
                        }
                    }

                    Console.WriteLine("Data uploaded successfully!");
                }
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"SQL Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General Error: {ex.Message}");
            }
        }
    }
}
