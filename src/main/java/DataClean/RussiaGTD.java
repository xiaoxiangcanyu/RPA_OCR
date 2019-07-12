package DataClean;

public class RussiaGTD {
    private String gtdNumber;
    private String gtdQuantity;
    private String gtdAmount;
    private String fileName;
    private String companyCode;
    private String status;

    public String getGtdNumber() {
        return gtdNumber;
    }

    public RussiaGTD setGtdNumber(String gtdNumber) {
        this.gtdNumber = gtdNumber;
        return this;
    }

    public String getGtdQuantity() {
        return gtdQuantity;
    }

    public RussiaGTD setGtdQuantity(String gtdQuantity) {
        this.gtdQuantity = gtdQuantity;
        return this;
    }

    public String getGtdAmount() {
        return gtdAmount;
    }

    public RussiaGTD setGtdAmount(String gtdAmount) {
        this.gtdAmount = gtdAmount;
        return this;
    }

    public String getFileName() {
        return fileName;
    }

    public RussiaGTD setFileName(String fileName) {
        this.fileName = fileName;
        return this;
    }

    public String getCompanyCode() {
        return companyCode;
    }

    public RussiaGTD setCompanyCode(String companyCode) {
        this.companyCode = companyCode;
        return this;
    }

    public String getStatus() {
        return status;
    }

    public RussiaGTD setStatus(String status) {
        this.status = status;
        return this;
    }
}
