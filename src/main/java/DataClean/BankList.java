package DataClean;

public class BankList {
    private String VendorName;
    private String accountNumber;
    private String sapCode;
    private String customerCode;
    private String taxCode;

    public String getVendorName() {
        return VendorName;
    }

    public void setVendorName(String vendorName) {
        VendorName = vendorName;
    }

    public String getAccountNumber() {
        return accountNumber;
    }

    public BankList setAccountNumber(String accountNumber) {
        this.accountNumber = accountNumber;
        return this;
    }

    public String getSapCode() {
        return sapCode;
    }

    public BankList setSapCode(String sapCode) {
        this.sapCode = sapCode;
        return this;
    }

    public String getCustomerCode() {
        return customerCode;
    }

    public void setCustomerCode(String customerCode) {
        this.customerCode = customerCode;
    }

    public String getTaxCode() {
        return taxCode;
    }

    public void setTaxCode(String taxCode) {
        this.taxCode = taxCode;
    }

    @Override
    public String toString() {
        return "BankList{" +
                "accountNumber='" + accountNumber + '\'' +
                ", sapCode='" + sapCode + '\'' +
                '}';
    }
}
