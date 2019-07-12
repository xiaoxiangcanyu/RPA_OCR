package DataClean;

public class Contrast {
    private String customerCode;
    private String reasonCode;
    private String type;
    private String text;

    public String getCustomerCode() {
        return customerCode;
    }

    public Contrast setCustomerCode(String customerCode) {
        this.customerCode = customerCode;
        return this;
    }

    public String getReasonCode() {
        return reasonCode;
    }

    public Contrast setReasonCode(String reasonCode) {
        this.reasonCode = reasonCode;
        return this;
    }

    public String getType() {
        return type;
    }

    public Contrast setType(String type) {
        this.type = type;
        return this;
    }

    public String getText() {
        return text;
    }

    public Contrast setText(String text) {
        this.text = text;
        return this;
    }

    @Override
    public String toString() {
        return "Contrast{" +
                "customerCode='" + customerCode + '\'' +
                ", reasonCode='" + reasonCode + '\'' +
                ", type='" + type + '\'' +
                ", text='" + text + '\'' +
                '}';
    }
}
