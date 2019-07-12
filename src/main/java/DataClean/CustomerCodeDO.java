package DataClean;

/**
 * 从customerCodeList里面读取的数据的实体类
 */
public class CustomerCodeDO {
    private String CustomerCode;
    private String Name;

    public String getCustomerCode() {
        return CustomerCode;
    }

    public void setCustomerCode(String customerCode) {
        CustomerCode = customerCode;
    }

    public String getName() {
        return Name;
    }

    public void setName(String name) {
        Name = name;
    }
}
