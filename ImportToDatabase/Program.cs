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
            string filePath = @"C:\Users\shahzaib.ahmed.DESKTOP-EB9K6GQ\Downloads\Address Book for PLP.xlsx";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var addressList = ReadExcelData(filePath);
            InsertDataToDatabase(addressList);

        }

        
        static List<AddressBook> ReadExcelData(string filePath)
        {
            var addresses = new List<AddressBook>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[4]; // Second worksheet
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    var address = new AddressBook();
                    address.Address = worksheet.Cells[row, 1].Text;
                    address.Address2 = worksheet.Cells[row, 2].Text;
                    address.City = worksheet.Cells[row, 3].Text;
                    address.State = worksheet.Cells[row, 4].Text;
                    address.Zip = (worksheet.Cells[row, 5].Text);
                    address.Country = (worksheet.Cells[row, 6].Text);
                    address.Phone = worksheet.Cells[row, 7].Text;
                    address.Fax = (worksheet.Cells[row, 8].Text);
                    address.PublicNotes = worksheet.Cells[row, 9].Text;
                    address.LoadPoints = worksheet.Cells[row, 10].Text;
                    address.Email = worksheet.Cells[row, 11].Text;
                    address.Contact = (worksheet.Cells[row, 12].Text);
                    address.Contact = worksheet.Cells[row, 13].Text;

                    addresses.Add(address);
                }
            }

            return addresses;
        }
        
        static void InsertDataToDatabase(List<AddressBook> addressList)
        {
            // Define the connection string
            string connectionString = "Server=tcp:brokerwareio20170123123330dbserver.database.windows.net,1433;Initial Catalog=3plbrokerwaredb_staging_2023-03-24T13-15Z;Persist Security Info=False; User ID=azuresa@brokerwareio20170123123330dbserver;Password=3pl$y$3760;MultipleActiveResultSets=True";
            //string customername = "Genetech1";
            //string customername = "UFPT";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    //int? customerId = GetCustomerId(connection, customername);
                    int? customerId = 4093902;

                    foreach (var address in addressList)
                    {
                        //company name outside


                        string query = @"
                                        INSERT INTO AddressBook (
                                            [Address]
                                           ,[Address2]
                                           ,[CellPhone]
                                           ,[City]
                                           ,[Country]

                                           ,[CreatedDate]
                                           ,[Fax]
                                           ,[ModifiedDate]
                                           ,[State]
                                           ,[WorkPhone]
                                           ,[Zip]
                                           
                                           ,[LoadPoint]
                                           
                                           ,[IsEditCityState])
                                        VALUES (
                                            @Address, @Address2, @CellPhone, @City, 
                                            @Country, @CreatedDate, @Fax, @ModifiedDate, 
                                            @State, @WorkPhone, @Zip, 
                                             @LoadPoint  , 1
                                        )
                                         SELECT SCOPE_IDENTITY();

";

                        int newAddressId;
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            // Add parameters to prevent SQL injection
                            command.Parameters.AddWithValue("@Address", address.Address);
                            command.Parameters.AddWithValue("@Address2", address.Address2 ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@CellPhone", address.Phone ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@City", address.City ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@Country", address.Country ?? (object)DBNull.Value); // Ensure Length is nullable
                            command.Parameters.AddWithValue("@CreatedDate", DateTime.UtcNow); // Ensure Width is nullable
                            command.Parameters.AddWithValue("@Fax", address.Fax); // Ensure Height is nullable
                            command.Parameters.AddWithValue("@ModifiedDate", DateTime.UtcNow);
                            command.Parameters.AddWithValue("@State", address.State); // Ensure Weight is nullable
                            command.Parameters.AddWithValue("@WorkPhone", address.Phone);
                            command.Parameters.AddWithValue("@Zip", address.Zip);
                            command.Parameters.AddWithValue("@LoadPoint", address.LoadPoints);


                            newAddressId = Convert.ToInt32(command.ExecuteScalar());
                        }

                        string insertCustomerAddressQuery = @"
                        INSERT INTO CustomerAddressBook (
                            [AddressId],
                            [CustomerId]
                        )
                        VALUES (
                            @AddressId,
                            @CustomerId
                        );";
                        using (SqlCommand command = new SqlCommand(insertCustomerAddressQuery, connection))
                        {
                            command.Parameters.AddWithValue("@AddressId", newAddressId);
                            command.Parameters.AddWithValue("@CustomerId", customerId);
                            command.ExecuteNonQuery();
                        }

                        Console.WriteLine("Data uploaded successfully!");
                    }
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
