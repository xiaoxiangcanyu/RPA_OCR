package DataClean;
/**
 * 各国发票数据对应实体类
 */
public class DataDO {
    private String Invoicenum;
    private String Amount;
    private String Invoicedate;
    private String InvoiceReferenceNumber;
    private String InvoiceReferenceNumber2;
    private String Poshorttext;
    private String PurchaseOrderNumber;
    private String Quantity;
    private String TaxAmount;
    private String TotalAmount;
    private String UnitPrice;
    private String CompanyCode;
    private String Currency;
    private String SOnumber;
    private String GoodDescription;
    private String PostingDate;
    private String TaxCode;
    private String  Status;
    private String Text;
    private String BaselineDate;
    private String ExchangeRate;
    private String PaymentBlock;
    private String Assignment;
    private String HeaderText;
    private String Filepath;
    private String OrderShipmentDate;
    private String ActualShipmentDate;
    private String OCRStatus;
    private String DownloadStatus;
    private String item;

    public String getInvoicenum() {
        return Invoicenum;
    }

    public void setInvoicenum(String invoicenum) {
        Invoicenum = invoicenum;
    }

    public String getAmount() {
        return Amount;
    }

    public void setAmount(String amount) {
        Amount = amount;
    }

    public String getInvoicedate() {
        return Invoicedate;
    }

    public void setInvoicedate(String invoicedate) {
        Invoicedate = invoicedate;
    }

    public String getInvoiceReferenceNumber() {
        return InvoiceReferenceNumber;
    }

    public void setInvoiceReferenceNumber(String invoiceReferenceNumber) {
        InvoiceReferenceNumber = invoiceReferenceNumber;
    }

    public String getInvoiceReferenceNumber2() {
        return InvoiceReferenceNumber2;
    }

    public void setInvoiceReferenceNumber2(String invoiceReferenceNumber2) {
        InvoiceReferenceNumber2 = invoiceReferenceNumber2;
    }

    public String getPoshorttext() {
        return Poshorttext;
    }

    public void setPoshorttext(String poshorttext) {
        Poshorttext = poshorttext;
    }

    public String getPurchaseOrderNumber() {
        return PurchaseOrderNumber;
    }

    public void setPurchaseOrderNumber(String purchaseOrderNumber) {
        PurchaseOrderNumber = purchaseOrderNumber;
    }

    public String getQuantity() {
        return Quantity;
    }

    public void setQuantity(String quantity) {
        Quantity = quantity;
    }

    public String getTaxAmount() {
        return TaxAmount;
    }

    public void setTaxAmount(String taxAmount) {
        TaxAmount = taxAmount;
    }

    public String getTotalAmount() {
        return TotalAmount;
    }

    public void setTotalAmount(String totalAmount) {
        TotalAmount = totalAmount;
    }

    public String getUnitPrice() {
        return UnitPrice;
    }

    public void setUnitPrice(String unitPrice) {
        UnitPrice = unitPrice;
    }

    public String getCompanyCode() {
        return CompanyCode;
    }

    public void setCompanyCode(String companyCode) {
        CompanyCode = companyCode;
    }

    public String getCurrency() {
        return Currency;
    }

    public void setCurrency(String currency) {
        Currency = currency;
    }

    public String getSOnumber() {
        return SOnumber;
    }

    public void setSOnumber(String SOnumber) {
        this.SOnumber = SOnumber;
    }

    public String getGoodDescription() {
        return GoodDescription;
    }

    public void setGoodDescription(String goodDescription) {
        GoodDescription = goodDescription;
    }

    public String getPostingDate() {
        return PostingDate;
    }

    public void setPostingDate(String postingDate) {
        PostingDate = postingDate;
    }

    public String getTaxCode() {
        return TaxCode;
    }

    public void setTaxCode(String taxCode) {
        TaxCode = taxCode;
    }

    public String getStatus() {
        return Status;
    }

    public void setStatus(String status) {
        Status = status;
    }

    public String getText() {
        return Text;
    }

    public void setText(String text) {
        Text = text;
    }

    public String getBaselineDate() {
        return BaselineDate;
    }

    public void setBaselineDate(String baselineDate) {
        BaselineDate = baselineDate;
    }

    public String getExchangeRate() {
        return ExchangeRate;
    }

    public void setExchangeRate(String exchangeRate) {
        ExchangeRate = exchangeRate;
    }

    public String getPaymentBlock() {
        return PaymentBlock;
    }

    public void setPaymentBlock(String paymentBlock) {
        PaymentBlock = paymentBlock;
    }

    public String getAssignment() {
        return Assignment;
    }

    public void setAssignment(String assignment) {
        Assignment = assignment;
    }

    public String getHeaderText() {
        return HeaderText;
    }

    public void setHeaderText(String headerText) {
        HeaderText = headerText;
    }

    public String getFilepath() {
        return Filepath;
    }

    public void setFilepath(String filepath) {
        Filepath = filepath;
    }

    public String getOrderShipmentDate() {
        return OrderShipmentDate;
    }

    public void setOrderShipmentDate(String orderShipmentDate) {
        OrderShipmentDate = orderShipmentDate;
    }

    public String getActualShipmentDate() {
        return ActualShipmentDate;
    }

    public void setActualShipmentDate(String actualShipmentDate) {
        ActualShipmentDate = actualShipmentDate;
    }

    public String getOCRStatus() {
        return OCRStatus;
    }

    public void setOCRStatus(String OCRStatus) {
        this.OCRStatus = OCRStatus;
    }

    public String getDownloadStatus() {
        return DownloadStatus;
    }

    public void setDownloadStatus(String downloadStatus) {
        DownloadStatus = downloadStatus;
    }

    public String getItem() {
        return item;
    }

    public DataDO setItem(String item) {
        this.item = item;
        return this;
    }

    @Override
    public String toString() {
        return "DataDO{" +
                "Invoicenum='" + Invoicenum + '\'' +
                ", Amount='" + Amount + '\'' +
                ", Invoicedate='" + Invoicedate + '\'' +
                ", InvoiceReferenceNumber='" + InvoiceReferenceNumber + '\'' +
                ", InvoiceReferenceNumber2='" + InvoiceReferenceNumber2 + '\'' +
                ", Poshorttext='" + Poshorttext + '\'' +
                ", PurchaseOrderNumber='" + PurchaseOrderNumber + '\'' +
                ", Quantity='" + Quantity + '\'' +
                ", TaxAmount='" + TaxAmount + '\'' +
                ", TotalAmount='" + TotalAmount + '\'' +
                ", UnitPrice='" + UnitPrice + '\'' +
                ", CompanyCode='" + CompanyCode + '\'' +
                ", Currency='" + Currency + '\'' +
                ", SOnumber='" + SOnumber + '\'' +
                ", GoodDescription='" + GoodDescription + '\'' +
                ", PostingDate='" + PostingDate + '\'' +
                ", TaxCode='" + TaxCode + '\'' +
                ", Status='" + Status + '\'' +
                ", Text='" + Text + '\'' +
                ", BaselineDate='" + BaselineDate + '\'' +
                ", ExchangeRate='" + ExchangeRate + '\'' +
                ", PaymentBlock='" + PaymentBlock + '\'' +
                ", Assignment='" + Assignment + '\'' +
                ", HeaderText='" + HeaderText + '\'' +
                ", Filepath='" + Filepath + '\'' +
                ", OrderShipmentDate='" + OrderShipmentDate + '\'' +
                ", ActualShipmentDate='" + ActualShipmentDate + '\'' +
                ", OCRStatus='" + OCRStatus + '\'' +
                ", DownloadStatus='" + DownloadStatus + '\'' +
                ", item='" + item + '\'' +
                '}';
    }
}
