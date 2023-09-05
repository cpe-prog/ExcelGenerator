using DocumentFormat.OpenXml.Office2016.Drawing.Command;
using ExcelGenerator;

var company = new Company()
{
    CompanyName = "DEBIT EXPRESS",
    Address = "Mansalay, Oriental Mindoro",
    Contact = 0935793759
};

var invoice = new Invoice()
{
    ITitle = "INVOICE",
    INumber = 8375,
    IDate = DateTime.Now,
    CostumerId = 764
};

var bill = new Bill()
{
    Name = "John Doe",
    CompanyName = "PCST",
    Address = "Mindoro",
    Phone = 03953957,
    Email = "sample@gmail.com"
};

var receipt = new Receipt()
{
    Description = "Sample Product", 
    Quantity = 2, 
    UnitPrice = 60
};

const string? filePath = @"C:\Users\grian\Desktop\Excel\Report.xlsx";
if (File.Exists(filePath))
{
    File.Delete(filePath);
}

SaveToExcell.Export(company, invoice, bill, receipt, filePath);
Console.WriteLine("Saved Successfully");
