package DataClean;

import java.util.List;

/**
 * 俄罗斯AR核销获取发票号对应实体类
 */
public class CustomerCode {
    private String taxCode;
    private String customerCode;
    private List<String> invoiceNoList;
    private String index;
    private String invoiceNo;
    private String tCode;
    private String country;
    private String taxNum;
    private String name;
    private String date;
    private String amount;
    private String text;

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

    public List<String> getInvoiceNoList() {
        return invoiceNoList;
    }

    public void setInvoiceNoList(List<String> invoiceNoList) {
        this.invoiceNoList = invoiceNoList;
    }

    public String getIndex() {
        return index;
    }

    public void setIndex(String index) {
        this.index = index;
    }

    public String getInvoiceNo() {
        return invoiceNo;
    }

    public void setInvoiceNo(String invoiceNo) {
        this.invoiceNo = invoiceNo;
    }

    public String gettCode() {
        return tCode;
    }

    public void settCode(String tCode) {
        this.tCode = tCode;
    }

    public String getCountry() {
        return country;
    }

    public void setCountry(String country) {
        this.country = country;
    }

    public String getTaxNum() {
        return taxNum;
    }

    public void setTaxNum(String taxNum) {
        this.taxNum = taxNum;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getAmount() {
        return amount;
    }

    public void setAmount(String amount) {
        this.amount = amount;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    @Override
    public String toString() {
        return "CustomerCode{" +
                "taxCode='" + taxCode + '\'' +
                ", customerCode='" + customerCode + '\'' +
                ", invoiceNoList=" + invoiceNoList +
                ", index='" + index + '\'' +
                ", invoiceNo='" + invoiceNo + '\'' +
                ", tCode='" + tCode + '\'' +
                ", country='" + country + '\'' +
                ", taxNum='" + taxNum + '\'' +
                ", name='" + name + '\'' +
                ", date='" + date + '\'' +
                ", amount='" + amount + '\'' +
                ", text='" + text + '\'' +
                '}';
    }
}
