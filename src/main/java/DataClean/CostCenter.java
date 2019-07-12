package DataClean;

public class CostCenter {
    private String productCode;
    private String costCenter;

    public String getProductCode() {
        return productCode;
    }

    public CostCenter setProductCode(String productCode) {
        this.productCode = productCode;
        return this;
    }

    public String getCostCenter() {
        return costCenter;
    }

    public CostCenter setCostCenter(String costCenter) {
        this.costCenter = costCenter;
        return this;
    }

    @Override
    public String toString() {
        return "CostCenter{" +
                "productCode='" + productCode + '\'' +
                ", costCenter='" + costCenter + '\'' +
                '}';
    }
}
