using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelGenerator;

public static class SaveToExcell
{
    public static void Export(Company companies, Invoice invoice, Bill bill,Receipt receipt,
        string filePath)
    {
        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Report");


        ws.ColumnWidth = 15;
        ws.Cell("A1").Value = companies.CompanyName;
        ws.Cell("A1").Style
            .Font.SetFontSize(15);
            ws.Cell("A1").Style.Font.Bold = true;
        
        ws.Cell("A2").Value = companies.Address;
        ws.Cell("A2").Style
            .Font.SetFontSize(12);
        
        ws.Cell("A3").Value = companies.Contact;
        ws.Cell("A3").Style
            .Font.SetFontSize(12);
        ws.Cell("A3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
        
        
        ws.Cell("E1").Value = invoice.ITitle;
        ws.Cell("D3").Value = "INVOICE #";
        ws.Cell("D4").Value = invoice.INumber;
        ws.Cell("D5").Value = "Customer ID";
        ws.Cell("D6").Value = invoice.CostumerId;
        ws.Cell("E3").Value = "Date";
        ws.Cell("E4").Value = invoice.IDate.Date;
        ws.Cell("E5").Value = "TERMS";
        ws.Cell("E6").Value = invoice.Terms;

        var rngTable = ws.Range("D3:E6");
        var borderRange = ws.Range("D3:D6");
        borderRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        rngTable.Style
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
            .Font.SetFontSize(12);
        
        ws.Cell("E1").Style
            .Font.SetFontColor(XLColor.Gray)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)
            .Font.SetFontSize(22)
            .Font.Bold = true;

        rngTable.Cell(1, 1).Style
            .Fill.SetBackgroundColor(XLColor.LightSlateGray)
            .Font.Bold = true;
        
        rngTable.Cell(1, 2).Style
            .Fill.SetBackgroundColor(XLColor.LightSlateGray)
            .Font.Bold = true;
       
        rngTable.Cell(3, 1).Style
            .Fill.SetBackgroundColor(XLColor.LightSlateGray)
            .Font.Bold = true;
        
        rngTable.Cell(3, 2).Style
            .Fill.SetBackgroundColor(XLColor.LightSlateGray)
            .Font.Bold = true;
        
        rngTable.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

        ws.Cell("A9").Value = "Bill";
        var billHeading = ws.Range("A9:B9");
        billHeading.Style
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
            .Fill.SetBackgroundColor(XLColor.Gray)
            .Font.SetFontSize(15)
            .Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Cell("A10").Value = bill.Name;
        ws.Cell("A11").Value = bill.CompanyName;
        ws.Cell("A12").Value = bill.Address;
        ws.Cell("A13").Value = bill.Phone;
        ws.Cell("A14").Value = bill.Email;
        var billRange = ws.Range("A9:A14");
        billRange.Style
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
            .Font.SetFontSize(12);

        ws.Cell("A17").Value = "Description";
        ws.Range("A17:B17").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        ws.Cell("C17").Value = "Quantity";
        ws.Cell("D17").Value = "UnitPrice";
        ws.Cell("E17").Value = "Description";
        var receiptHeadingRange = ws.Range("A17:E17");
        receiptHeadingRange.Style
            .Fill.SetBackgroundColor(XLColor.Gray)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
            .Font.SetFontSize(12)
            .Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Cell("A18").Value = receipt.Description;
        ws.Cell("C18").Value = receipt.Quantity;
        ws.Cell("D18").Value = receipt.UnitPrice;
        ws.Cell("E18").Value = receipt.Amount;
        ws.Range("A18:A38").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
        var receiptRange = ws.Range("A18:E18");
        var brRange = ws.Range("C17:E38");
        receiptRange.Style
            .Font.SetFontSize(12);
        brRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        
        var receiptTable = ws.Range("A17:E40");
        receiptTable.Style
            .Font.SetFontSize(12)
            .Border.OutsideBorder = XLBorderStyleValues.Thin;

        ws.Cell("C40").Value = "TOTAL";
        ws.Cell("E40").Value = "0000";
        ws.Cell("E40").Style
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
        var topBorder = ws.Range("A39:E40");
        var leftBorder = ws.Range("C40");
        topBorder.Style
            .Border.SetTopBorder(XLBorderStyleValues.Thin);
        leftBorder.Style
            .Border.SetLeftBorder(XLBorderStyleValues.Thin);
        
        
        workbook.SaveAs(filePath);
    }
}