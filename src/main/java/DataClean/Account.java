package DataClean;

public class Account {
    private String bank;
    private String bankCurrency;
    private String currency;
    private String account;

    public String getBank() {
        return bank;
    }

    public Account setBank(String bank) {
        this.bank = bank;
        return this;
    }

    public String getBankCurrency() {
        return bankCurrency;
    }

    public Account setBankCurrency(String bankCurrency) {
        this.bankCurrency = bankCurrency;
        return this;
    }

    public String getCurrency() {
        return currency;
    }

    public Account setCurrency(String currency) {
        this.currency = currency;
        return this;
    }

    public String getAccount() {
        return account;
    }

    public Account setAccount(String account) {
        this.account = account;
        return this;
    }


    @Override
    public String toString() {
        return "Account{" +
                "bank='" + bank + '\'' +
                ", bankCurrency='" + bankCurrency + '\'' +
                ", currency='" + currency + '\'' +
                ", account='" + account + '\'' +
                '}';
    }
}
