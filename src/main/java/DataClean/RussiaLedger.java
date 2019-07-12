package DataClean;

public class RussiaLedger {
    private String no;
    private String gl;
    private String date;
    private String vendor;
    private String title;
    private String vendorCode;
    private String amount;
    private String reasonCode;
    private String sapIncomeNo;
    private String specialGL;
    private String invoice;
    private String sapClearNo;
    private String status;
    private String tax;
    private String taxCode;
    private String text;

    public String getNo() {
        return no;
    }

    public RussiaLedger setNo(String no) {
        this.no = no;
        return this;
    }

    public String getGl() {
        return gl;
    }

    public RussiaLedger setGl(String gl) {
        this.gl = gl;
        return this;
    }

    public String getDate() {
        return date;
    }

    public RussiaLedger setDate(String date) {
        this.date = date;
        return this;
    }

    public String getVendor() {
        return vendor;
    }

    public RussiaLedger setVendor(String vendor) {
        this.vendor = vendor;
        return this;
    }

    public String getTitle() {
        return title;
    }

    public RussiaLedger setTitle(String title) {
        this.title = title;
        return this;
    }

    public String getVendorCode() {
        return vendorCode;
    }

    public RussiaLedger setVendorCode(String vendorCode) {
        this.vendorCode = vendorCode;
        return this;
    }

    public String getAmount() {
        return amount;
    }

    public RussiaLedger setAmount(String amount) {
        this.amount = amount;
        return this;
    }

    public String getReasonCode() {
        return reasonCode;
    }

    public RussiaLedger setReasonCode(String reasonCode) {
        this.reasonCode = reasonCode;
        return this;
    }

    public String getSapIncomeNo() {
        return sapIncomeNo;
    }

    public RussiaLedger setSapIncomeNo(String sapIncomeNo) {
        this.sapIncomeNo = sapIncomeNo;
        return this;
    }

    public String getSpecialGL() {
        return specialGL;
    }

    public RussiaLedger setSpecialGL(String specialGL) {
        this.specialGL = specialGL;
        return this;
    }

    public String getInvoice() {
        return invoice;
    }

    public RussiaLedger setInvoice(String invoice) {
        this.invoice = invoice;
        return this;
    }

    public String getSapClearNo() {
        return sapClearNo;
    }

    public RussiaLedger setSapClearNo(String sapClearNo) {
        this.sapClearNo = sapClearNo;
        return this;
    }

    public String getStatus() {
        return status;
    }

    public RussiaLedger setStatus(String status) {
        this.status = status;
        return this;
    }

    public String getTaxCode() {
        return taxCode;
    }

    public RussiaLedger setTaxCode(String taxCode) {
        this.taxCode = taxCode;
        return this;
    }

    public String getText() {
        return text;
    }

    public RussiaLedger setText(String text) {
        this.text = text;
        return this;
    }

    public String getTax() {
        return tax;
    }

    public void setTax(String tax) {
        this.tax = tax;
    }

    @Override
    public String toString() {
        return "RussiaLedger{" +
                "no='" + no + '\'' +
                ", gl='" + gl + '\'' +
                ", date='" + date + '\'' +
                ", vendor='" + vendor + '\'' +
                ", title='" + title + '\'' +
                ", vendorCode='" + vendorCode + '\'' +
                ", amount='" + amount + '\'' +
                ", reasonCode='" + reasonCode + '\'' +
                ", sapIncomeNo='" + sapIncomeNo + '\'' +
                ", specialGL='" + specialGL + '\'' +
                ", invoice='" + invoice + '\'' +
                ", sapClearNo='" + sapClearNo + '\'' +
                ", status='" + status + '\'' +
                ", taxCode='" + taxCode + '\'' +
                ", text='" + text + '\'' +
                '}';
    }
}
