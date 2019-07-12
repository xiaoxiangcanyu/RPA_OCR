package DataClean;
/**
 * 俄罗斯AR核销对应实体类
 */
public class Ledger {
    private String date;
    private String customer;
    private String title;
    private String taxCode;
    private String customerCode;
    private String amount;
    private String reasonCode;
    private String invoice;
    private String sapIncomeNo;
    private String sapClearNo;
    private String status;
    private String receivedInforDate;
    private String keyinDate;
    private String comment;
    private String text;
    private String country;
    private String GL;

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getCustomer() {
        return customer;
    }

    public void setCustomer(String customer) {
        this.customer = customer;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getTaxCode() {
        return taxCode;
    }

    public void setTaxCode(String taxCode) {
        this.taxCode = taxCode;
    }

    public String getCustomerCode() {
        return customerCode;
    }

    public void setCustomerCode(String customerCode) {
        this.customerCode = customerCode;
    }

    public String getAmount() {
        return amount;
    }

    public void setAmount(String amount) {
        this.amount = amount;
    }

    public String getReasonCode() {
        return reasonCode;
    }

    public void setReasonCode(String reasonCode) {
        this.reasonCode = reasonCode;
    }

    public String getInvoice() {
        return invoice;
    }

    public void setInvoice(String invoice) {
        this.invoice = invoice;
    }

    public String getSapIncomeNo() {
        return sapIncomeNo;
    }

    public void setSapIncomeNo(String sapIncomeNo) {
        this.sapIncomeNo = sapIncomeNo;
    }

    public String getSapClearNo() {
        return sapClearNo;
    }

    public void setSapClearNo(String sapClearNo) {
        this.sapClearNo = sapClearNo;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public String getReceivedInforDate() {
        return receivedInforDate;
    }

    public void setReceivedInforDate(String receivedInforDate) {
        this.receivedInforDate = receivedInforDate;
    }

    public String getKeyinDate() {
        return keyinDate;
    }

    public void setKeyinDate(String keyinDate) {
        this.keyinDate = keyinDate;
    }

    public String getComment() {
        return comment;
    }

    public void setComment(String comment) {
        this.comment = comment;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public String getCountry() {
        return country;
    }

    public void setCountry(String country) {
        this.country = country;
    }

    public String getGL() {
        return GL;
    }

    public void setGL(String GL) {
        this.GL = GL;
    }

    @Override
    public String toString() {
        return "Ledger{" +
                "date='" + date + '\'' +
                ", customer='" + customer + '\'' +
                ", title='" + title + '\'' +
                ", taxCode='" + taxCode + '\'' +
                ", customerCode='" + customerCode + '\'' +
                ", amount='" + amount + '\'' +
                ", reasonCode='" + reasonCode + '\'' +
                ", invoice='" + invoice + '\'' +
                ", sapIncomeNo='" + sapIncomeNo + '\'' +
                ", sapClearNo='" + sapClearNo + '\'' +
                ", status='" + status + '\'' +
                ", receivedInforDate='" + receivedInforDate + '\'' +
                ", keyinDate='" + keyinDate + '\'' +
                ", comment='" + comment + '\'' +
                ", text='" + text + '\'' +
                ", country='" + country + '\'' +
                '}';
    }
}
