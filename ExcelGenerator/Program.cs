﻿using ClosedXML;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelGenerator;

var company = new Company()
{
    CompanyName = "Microsoft",
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
    CompanyName = "FaceBook",
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
SaveToExcell.Export(company, invoice, bill, receipt, filePath);

Console.WriteLine("Saved Successfully");