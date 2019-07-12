package DataClean;

public class MiddleEast {
    private String no;//序号
    private String bank;//收款银行
    private String documentDate;//收款日期
    private String amount;//收款金额
    private String currency;//币种
    private String summary;//银行附言
    private String customerName;//客户名称
    private String customerCode;//客户代码
    private String recognizedAmount;//认领金额
    private String pi;//
    private String productCode;//产品线
    private String invoiceNumber;//发票号
    private String prepayment;//预付比例
    private String stafName;//认领人
    private String contact;//电话或邮箱
    private String note;//备注
    private String state;//数据状态
    private String costCentre;
    private String id;
    private String sapIncomeNo;
    private String sapClearNo;
    private String status;
    private String receivedInforDate;
    private String keyInDate;
    private String comment;

    public String getNo() {
        return no;
    }

    public void setNo(String no) {
        this.no = no;
    }

    public String getBank() {
        return bank;
    }

    public void setBank(String bank) {
        this.bank = bank;
    }

    public String getDocumentDate() {
        return documentDate;
    }

    public void setDocumentDate(String documentDate) {
        this.documentDate = documentDate;
    }

    public String getAmount() {
        return amount;
    }

    public void setAmount(String amount) {
        this.amount = amount;
    }

    public String getCurrency() {
        return currency;
    }

    public void setCurrency(String currency) {
        this.currency = currency;
    }

    public String getSummary() {
        return summary;
    }

    public void setSummary(String summary) {
        this.summary = summary;
    }

    public String getCustomerName() {
        return customerName;
    }

    public void setCustomerName(String customerName) {
        this.customerName = customerName;
    }

    public String getCustomerCode() {
        return customerCode;
    }

    public void setCustomerCode(String customerCode) {
        this.customerCode = customerCode;
    }

    public String getRecognizedAmount() {
        return recognizedAmount;
    }

    public void setRecognizedAmount(String recognizedAmount) {
        this.recognizedAmount = recognizedAmount;
    }

    public String getPi() {
        return pi;
    }

    public void setPi(String pi) {
        this.pi = pi;
    }

    public String getProductCode() {
        return productCode;
    }

    public void setProductCode(String productCode) {
        this.productCode = productCode;
    }

    public String getInvoiceNumber() {
        return invoiceNumber;
    }

    public void setInvoiceNumber(String invoiceNumber) {
        this.invoiceNumber = invoiceNumber;
    }

    public String getPrepayment() {
        return prepayment;
    }

    public void setPrepayment(String prepayment) {
        this.prepayment = prepayment;
    }

    public String getStafName() {
        return stafName;
    }

    public void setStafName(String stafName) {
        this.stafName = stafName;
    }

    public String getContact() {
        return contact;
    }

    public void setContact(String contact) {
        this.contact = contact;
    }

    public String getNote() {
        return note;
    }

    public void setNote(String note) {
        this.note = note;
    }

    public String getStatus() {
        return status;
    }

    public MiddleEast setStatus(String status) {
        this.status = status;
        return this;
    }

    public String getCostCentre() {
        return costCentre;
    }

    public MiddleEast setCostCentre(String costCentre) {
        this.costCentre = costCentre;
        return this;
    }

    public String getId() {
        return id;
    }

    public MiddleEast setId(String id) {
        this.id = id;
        return this;
    }

    public String getSapIncomeNo() {
        return sapIncomeNo;
    }

    public MiddleEast setSapIncomeNo(String sapIncomeNo) {
        this.sapIncomeNo = sapIncomeNo;
        return this;
    }

    public String getState() {
        return state;
    }

    public MiddleEast setState(String state) {
        this.state = state;
        return this;
    }

    public String getSapClearNo() {
        return sapClearNo;
    }

    public MiddleEast setSapClearNo(String sapClearNo) {
        this.sapClearNo = sapClearNo;
        return this;
    }

    public String getReceivedInforDate() {
        return receivedInforDate;
    }

    public MiddleEast setReceivedInforDate(String receivedInforDate) {
        this.receivedInforDate = receivedInforDate;
        return this;
    }

    public String getKeyInDate() {
        return keyInDate;
    }

    public MiddleEast setKeyInDate(String keyInDate) {
        this.keyInDate = keyInDate;
        return this;
    }

    public String getComment() {
        return comment;
    }

    public MiddleEast setComment(String comment) {
        this.comment = comment;
        return this;
    }

    @Override
    public String toString() {
        return "MiddleEast{" +
                "no='" + no + '\'' +
                ", bank='" + bank + '\'' +
                ", documentDate='" + documentDate + '\'' +
                ", amount='" + amount + '\'' +
                ", currency='" + currency + '\'' +
                ", summary='" + summary + '\'' +
                ", customerName='" + customerName + '\'' +
                ", customerCode='" + customerCode + '\'' +
                ", recognizedAmount='" + recognizedAmount + '\'' +
                ", pi='" + pi + '\'' +
                ", productCode='" + productCode + '\'' +
                ", invoiceNumber='" + invoiceNumber + '\'' +
                ", prepayment='" + prepayment + '\'' +
                ", stafName='" + stafName + '\'' +
                ", contact='" + contact + '\'' +
                ", note='" + note + '\'' +
                ", state='" + state + '\'' +
                ", costCentre='" + costCentre + '\'' +
                ", id='" + id + '\'' +
                ", sapIncomeNo='" + sapIncomeNo + '\'' +
                ", sapClearNo='" + sapClearNo + '\'' +
                ", status='" + status + '\'' +
                ", receivedInforDate='" + receivedInforDate + '\'' +
                ", keyInDate='" + keyInDate + '\'' +
                ", comment='" + comment + '\'' +
                '}';
    }
}
