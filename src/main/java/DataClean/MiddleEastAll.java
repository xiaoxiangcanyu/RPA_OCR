package DataClean;

public class MiddleEastAll {
    private String id;
    private String no;
    private String bank;
    private String documentDate;
    private String income;//收款
    private String charge;//扣款
    private String currency;//货币号
    private String summary;//银行附言
    private String ttLCMark;//TT/LC 类型数据标识
    private String customerName;//客户名称
    private String customerCode;//客户编码
    private String recognizedAmount;//认领金额
    private String productCode;//产品线
    private String pi;
    private String prepayment;
    private String transferTo;
    private String depositPrincipal;
    private String depositInterest;
    private String staffName;
    private String remark;
    private String sapIncomeNo;//状态标识
    private String comment;
    private String balance;//余额
    private String sapNoForOther;//一次性客户标识
    private String sapClearingNo;
    private String email;//邮箱

    public String getId() {
        return id;
    }

    public MiddleEastAll setId(String id) {
        this.id = id;
        return this;
    }

    public String getNo() {
        return no;
    }

    public MiddleEastAll setNo(String no) {
        this.no = no;
        return this;
    }

    public String getBank() {
        return bank;
    }

    public MiddleEastAll setBank(String bank) {
        this.bank = bank;
        return this;
    }

    public String getDocumentDate() {
        return documentDate;
    }

    public MiddleEastAll setDocumentDate(String documentDate) {
        this.documentDate = documentDate;
        return this;
    }

    public String getIncome() {
        return income;
    }

    public MiddleEastAll setIncome(String income) {
        this.income = income;
        return this;
    }

    public String getCharge() {
        return charge;
    }

    public MiddleEastAll setCharge(String charge) {
        this.charge = charge;
        return this;
    }

    public String getCurrency() {
        return currency;
    }

    public MiddleEastAll setCurrency(String currency) {
        this.currency = currency;
        return this;
    }

    public String getSummary() {
        return summary;
    }

    public MiddleEastAll setSummary(String summary) {
        this.summary = summary;
        return this;
    }

    public String getTtLCMark() {
        return ttLCMark;
    }

    public MiddleEastAll setTtLCMark(String ttLCMark) {
        this.ttLCMark = ttLCMark;
        return this;
    }

    public String getCustomerName() {
        return customerName;
    }

    public MiddleEastAll setCustomerName(String customerName) {
        this.customerName = customerName;
        return this;
    }

    public String getCustomerCode() {
        return customerCode;
    }

    public MiddleEastAll setCustomerCode(String customerCode) {
        this.customerCode = customerCode;
        return this;
    }

    public String getRecognizedAmount() {
        return recognizedAmount;
    }

    public MiddleEastAll setRecognizedAmount(String recognizedAmount) {
        this.recognizedAmount = recognizedAmount;
        return this;
    }

    public String getProductCode() {
        return productCode;
    }

    public MiddleEastAll setProductCode(String productCode) {
        this.productCode = productCode;
        return this;
    }

    public String getPi() {
        return pi;
    }

    public MiddleEastAll setPi(String pi) {
        this.pi = pi;
        return this;
    }

    public String getPrepayment() {
        return prepayment;
    }

    public MiddleEastAll setPrepayment(String prepayment) {
        this.prepayment = prepayment;
        return this;
    }

    public String getTransferTo() {
        return transferTo;
    }

    public MiddleEastAll setTransferTo(String transferTo) {
        this.transferTo = transferTo;
        return this;
    }

    public String getDepositPrincipal() {
        return depositPrincipal;
    }

    public MiddleEastAll setDepositPrincipal(String depositPrincipal) {
        this.depositPrincipal = depositPrincipal;
        return this;
    }

    public String getDepositInterest() {
        return depositInterest;
    }

    public MiddleEastAll setDepositInterest(String depositInterest) {
        this.depositInterest = depositInterest;
        return this;
    }

    public String getStaffName() {
        return staffName;
    }

    public MiddleEastAll setStaffName(String staffName) {
        this.staffName = staffName;
        return this;
    }

    public String getRemark() {
        return remark;
    }

    public MiddleEastAll setRemark(String remark) {
        this.remark = remark;
        return this;
    }

    public String getSapIncomeNo() {
        return sapIncomeNo;
    }

    public MiddleEastAll setSapIncomeNo(String sapIncomeNo) {
        this.sapIncomeNo = sapIncomeNo;
        return this;
    }

    public String getComment() {
        return comment;
    }

    public MiddleEastAll setComment(String comment) {
        this.comment = comment;
        return this;
    }

    public String getBalance() {
        return balance;
    }

    public MiddleEastAll setBalance(String balance) {
        this.balance = balance;
        return this;
    }

    public String getSapNoForOther() {
        return sapNoForOther;
    }

    public MiddleEastAll setSapNoForOther(String sapNoForOther) {
        this.sapNoForOther = sapNoForOther;
        return this;
    }

    public String getSapClearingNo() {
        return sapClearingNo;
    }

    public MiddleEastAll setSapClearingNo(String sapClearingNo) {
        this.sapClearingNo = sapClearingNo;
        return this;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    @Override
    public String toString() {
        return "MiddleEastAll{" +
                "id='" + id + '\'' +
                ", no='" + no + '\'' +
                ", bank='" + bank + '\'' +
                ", documentDate='" + documentDate + '\'' +
                ", income='" + income + '\'' +
                ", charge='" + charge + '\'' +
                ", currency='" + currency + '\'' +
                ", summary='" + summary + '\'' +
                ", ttLCMark='" + ttLCMark + '\'' +
                ", customerName='" + customerName + '\'' +
                ", customerCode='" + customerCode + '\'' +
                ", recognizedAmount='" + recognizedAmount + '\'' +
                ", productCode='" + productCode + '\'' +
                ", pi='" + pi + '\'' +
                ", prepayment='" + prepayment + '\'' +
                ", transferTo='" + transferTo + '\'' +
                ", depositPrincipal='" + depositPrincipal + '\'' +
                ", depositInterest='" + depositInterest + '\'' +
                ", staffName='" + staffName + '\'' +
                ", remark='" + remark + '\'' +
                ", sapIncomeNo='" + sapIncomeNo + '\'' +
                ", comment='" + comment + '\'' +
                ", balance='" + balance + '\'' +
                ", sapNoForOther='" + sapNoForOther + '\'' +
                ", sapClearingNo='" + sapClearingNo + '\'' +
                '}';
    }
}
