using ClosedXML.Excel;

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
                .Font.SetFontSize(15)
                .Font.SetFontName("Arial Narrow");
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.SetFontColor(XLColor.Blue);
        
        ws.Cell("A2").Value = companies.Address;
        ws.Cell("A2").Style
            .Font.SetFontSize(12)
            .Font.SetFontName("Arial Narrow");
        
        ws.Cell("A3").Value = companies.Contact;
        ws.Cell("A3").Style
            .Font.SetFontSize(12)
            .Font.SetFontName("Arial Narrow");
        ws.Cell("A3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

        ws.Cell("A5").Value = invoice.ITitle;
        ws.Cell("A6").Value = invoice.CostumerId;
        ws.Cell("A7").Value = invoice.IDate;
        ws.Cell("A8").Value = invoice.INumber;

        ws.Cell("A10").Value = bill.Name;
        ws.Cell("A11").Value = bill.CompanyName;
        ws.Cell("A12").Value = bill.Address;
        ws.Cell("A13").Value = bill.Phone;
        ws.Cell("A14").Value = bill.Email;

        ws.Cell("A16").Value = receipt.Description;
        ws.Cell("A17").Value = receipt.Quantity;
        ws.Cell("A18").Value = receipt.UnitPrice;
        ws.Cell("A18").Value = receipt.UnitPrice;
        ws.Cell("B19").Value = receipt.Amount;

        workbook.SaveAs(filePath);
    }
}